using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;

namespace RackPowerUPS
{
    public class ModbusFrame
    {
        public byte SlaveAddress { get; set; }
        public byte FunctionCode { get; set; }
        public byte[] Data { get; set; } = new byte[0];
        public ushort Crc { get; set; }
        public bool CrcValid { get; set; }
        public bool IsException => (FunctionCode & 0x80) != 0;
    }

    public static class ModbusHelper
    {
        public static ModbusFrame ParseFrame(byte[] raw)
        {
            if (raw == null || raw.Length < 4)
                throw new ArgumentException("Frame too short", nameof(raw));

            ushort rawCrc = BitConverter.ToUInt16(raw, raw.Length - 2);
            ushort calc = Crc16(raw, raw.Length - 2);

            return new ModbusFrame
            {
                SlaveAddress = raw[0],
                FunctionCode = raw[1],
                Data = raw.Skip(2).Take(raw.Length - 4).ToArray(),
                Crc = rawCrc,
                CrcValid = rawCrc == calc
            };
        }

        public static ushort[] ExtractRegisters(ModbusFrame frame)
        {
            if (frame == null) throw new ArgumentNullException(nameof(frame));
            if (frame.Data == null || frame.Data.Length < 2) return new ushort[0];

            int byteCount = frame.Data[0];
            int offset = (byteCount == frame.Data.Length - 1) ? 1 : 0;
            int regCount = (frame.Data.Length - offset) / 2;

            ushort[] regs = new ushort[regCount];
            for (int i = 0; i < regCount; i++)
            {
                int idx = offset + i * 2;
                regs[i] = (ushort)((frame.Data[idx] << 8) | frame.Data[idx + 1]);
            }
            return regs;
        }

        public static ushort Crc16(byte[] data, int length = -1)
        {
            if (data == null) return 0;
            if (length < 0) length = data.Length;

            ushort crc = 0xFFFF;
            for (int pos = 0; pos < length; pos++)
            {
                crc ^= data[pos];
                for (int i = 0; i < 8; i++)
                {
                    bool lsb = (crc & 0x0001) != 0;
                    crc >>= 1;
                    if (lsb) crc ^= 0xA001;
                }
            }
            return crc;
        }

        public static string ByteToString(byte[] regs, int regIndex, int lengthRegs)
        {
            if (regs == null) { throw new ArgumentNullException(nameof(regs)); }
            if (regIndex < 0 || regIndex + lengthRegs > regs.Length) { throw new IndexOutOfRangeException(); }
            return Encoding.ASCII.GetString(regs, regIndex, lengthRegs);
        }
    }

    public class RackPowerUpsClient : IDisposable
    {
        public SerialPort _serialPort;
        public ushort[] _regs;

        private readonly object _sync = new object();
        private readonly Dictionary<string, DateTime> _lastUpdated = new Dictionary<string, DateTime>();
        public TimeSpan _maxAge = TimeSpan.FromSeconds(1);

        public enum SwitchState
        {
            BatteryNotConnected = 0,
            Normal = 1,
            MBCBClosed_BatNotConnected = 2,
            MBCBClosed = 3,
            EPO_BatteryNotConnected = 4,
            EPO = 5,
            MBCBClosed_BatNotConnected_EPO = 6,
            MBCBClosed_EPO = 7
        }

        public enum OutputStatus
        {
            NoOutput = 0,
            Inverter = 1,
            Diode = 2
        }

        public enum HardwareType
        {
            Unknown = 0,
            ACM_10_600kVA = 1,
            TT_10_40kVA = 2,
            TS_10_20kVA = 3,
            SS_6_20kVA = 4,
            SS_1_3kVA = 5
        }

        [Flags]
        public enum SystemStatusFlags
        {
            None = 0,
            AmbientOverTemp = 1 << 0, // 0x1
            RecCanFail = 1 << 1, // 0x2
            InvIoCanFail = 1 << 2, // 0x4
            InvDataCanFail = 1 << 3  // 0x8
        }

        [Flags]
        public enum SystemStatus2Flags
        {
            None = 0,
            BypassPowerFuseFail = 1 << 0, // 0x1
            RatedKvaOverRange = 1 << 1  // 0x2
        }

        [Flags]
        public enum SystemStatus3Flags
        {
            None = 0,
            No_IP_SCR_TmpSensor = 1 << 1, // 0x2
            IP_SCR_Over_Temp = 1 << 2    // 0x4
        }

        // UPS Values
        // === Cached Properties with EnsureFresh ===
        private string _model;
        public string Model => EnsureFresh("ManufacturerInfo", () => _model, QueryManufacturerInfo, runOnce: true);

        private string _version;
        public string Version => EnsureFresh("VersionTemps", () => _version, QueryVersionTemps);

        private int _batteryCount;
        public int BatteryCount => EnsureFresh("ManufacturerInfo", () => _batteryCount, QueryManufacturerInfo, runOnce: true);

        private float _bypassVoltage;
        public float BypassVoltage => EnsureFresh("PowerInfo", () => _bypassVoltage, QueryPowerInfo);

        private float _bypassCurrent;
        public float BypassCurrent => EnsureFresh("PowerInfo", () => _bypassCurrent, QueryPowerInfo);

        private float _mainVoltage;
        public float MainVoltage => EnsureFresh("PowerInfo", () => _mainVoltage, QueryPowerInfo);

        private float _mainVoltage2;
        public float MainVoltage2 => EnsureFresh("PowerInfo", () => _mainVoltage2, QueryPowerInfo);

        private float _mainVoltage3;
        public float MainVoltage3 => EnsureFresh("PowerInfo", () => _mainVoltage3, QueryPowerInfo);

        private float _outputVoltage;
        public float OutputVoltage => EnsureFresh("PowerInfo", () => _outputVoltage, QueryPowerInfo);

        private float _va;
        public float VA => EnsureFresh("PowerInfo", () => _va, QueryPowerInfo);

        private float _outputReactivePower;
        public float OutputReactivePower => EnsureFresh("PowerInfo", () => _outputReactivePower, QueryPowerInfo);

        private float _watts;
        public float Watts => EnsureFresh("PowerInfo", () => _watts, QueryPowerInfo);

        private float _batteryVoltage;
        public float BatteryVoltage => EnsureFresh("PowerInfo", () => _batteryVoltage, QueryPowerInfo);

        private float _batteryCurrent;
        public float BatteryCurrent => EnsureFresh("PowerInfo", () => _batteryCurrent, QueryPowerInfo);

        private float _batteryCapacity;
        public float BatteryCapacity => EnsureFresh("PowerInfo", () => _batteryCapacity, QueryPowerInfo);

        private float _batteryTimeRemain;
        public float BatteryTimeRemain => EnsureFresh("PowerInfo", () => _batteryTimeRemain, QueryPowerInfo);

        private float _batteryTemp;
        public float BatteryTemp => EnsureFresh("PowerInfo", () => _batteryTemp, QueryPowerInfo);

        private float _runtime;
        public float Runtime => EnsureFresh("PowerInfo", () => _runtime, QueryPowerInfo);

        private float _bypassFanHour;
        public float BypassFanHour => EnsureFresh("PowerInfo", () => _bypassFanHour, QueryPowerInfo);

        private float _dustFilterDays;
        public float DustFilterDays => EnsureFresh("PowerInfo", () => _dustFilterDays, QueryPowerInfo);

        private float _dischargehours;
        public float DischargeHours => EnsureFresh("PowerInfo", () => _dischargehours, QueryPowerInfo);

        private float _dischargeamount;
        public float DischargeAmount => EnsureFresh("PowerInfo", () => _dischargeamount, QueryPowerInfo);

        private float _lockremaintime;
        public float LockRemainTime => EnsureFresh("PowerInfo", () => _lockremaintime, QueryPowerInfo);

        private float _upskey;
        public float UPSKey => EnsureFresh("PowerInfo", () => _upskey, QueryPowerInfo);

        private float _upstype;
        public float UPSType => EnsureFresh("PowerInfo", () => _upstype, QueryPowerInfo);

        private float _currentcorrectablemark;
        public float CurrentCorrectableMark => EnsureFresh("PowerInfo", () => _currentcorrectablemark, QueryPowerInfo);

        private float _inverteravgvolt;
        public float InverterAvgVolt => EnsureFresh("PowerInfo", () => _inverteravgvolt, QueryPowerInfo);

        private float _bypassavgvolt;
        public float BypassAvgVolt => EnsureFresh("PowerInfo", () => _bypassavgvolt, QueryPowerInfo);

        private float _selfagingmode;
        public float SelfAgingMode => EnsureFresh("PowerInfo", () => _selfagingmode, QueryPowerInfo);

        private float _busVoltage;
        public float BusVoltage => EnsureFresh("PowerInfo", () => _busVoltage, QueryPowerInfo);

        private float _syscode;
        public float Syscode => EnsureFresh("PowerInfo", () => _syscode, QueryPowerInfo);

        private float _ambientTemp;
        public float AmbientTemp => EnsureFresh("PowerInfo", () => _ambientTemp, QueryPowerInfo);

        private float _loadPercent;
        public float LoadPercent => EnsureFresh("PowerInfo", () => _loadPercent, QueryPowerInfo);

        private float _bypassFreq;
        public float BypassFreq => EnsureFresh("PowerInfo", () => _bypassFreq, QueryPowerInfo);

        private float _mainFreq;
        public float MainFreq => EnsureFresh("PowerInfo", () => _mainFreq, QueryPowerInfo);

        private float _mainFreq2;
        public float MainFreq2 => EnsureFresh("PowerInfo", () => _mainFreq2, QueryPowerInfo);

        private float _mainFreq3;
        public float MainFreq3 => EnsureFresh("PowerInfo", () => _mainFreq3, QueryPowerInfo);

        private float _outputFreq;
        public float OutputFreq => EnsureFresh("PowerInfo", () => _outputFreq, QueryPowerInfo);

        private float _ratedOutputVoltage;
        public float RatedOutputVoltage => EnsureFresh("PowerInfo", () => _ratedOutputVoltage, QueryPowerInfo);

        private float _ratedOutputFreq;
        public float RatedOutputFreq => EnsureFresh("PowerInfo", () => _ratedOutputFreq, QueryPowerInfo);

        private float _ratedInputVoltage;
        public float RatedInputVoltage => EnsureFresh("PowerInfo", () => _ratedInputVoltage, QueryPowerInfo);

        private float _ratedInputFreq;
        public float RatedInputFreq => EnsureFresh("PowerInfo", () => _ratedInputFreq, QueryPowerInfo);

        private float _chargerCurrent;
        public float ChargerCurrent => EnsureFresh("PowerInfo", () => _chargerCurrent, QueryPowerInfo);

        private float _chargerVoltage;
        public float ChargerVoltage => EnsureFresh("PowerInfo", () => _chargerVoltage, QueryPowerInfo);

        private float _inverterVoltage;
        public float InverterVoltage => EnsureFresh("PowerInfo", () => _inverterVoltage, QueryPowerInfo);

        private float _fanruntime;
        public float FanRunTime => EnsureFresh("PowerInfo", () => _fanruntime, QueryPowerInfo);

        private float _capacitortime;
        public float CapacitorTime => EnsureFresh("PowerInfo", () => _capacitortime, QueryPowerInfo);

        private float _recIgbtTemp;
        public float RecIGBTTemp => EnsureFresh("VersionTemps", () => _recIgbtTemp, QueryVersionTemps);

        private float _invIgbtTemp;
        public float InvIGBTTemp => EnsureFresh("VersionTemps", () => _invIgbtTemp, QueryVersionTemps);

        private float _maxWatt;
        public float MaxWatt => EnsureFresh("PowerInfo", () => _maxWatt, QueryPowerInfo);

        private float _mainCurrent;
        public float MainCurrent => EnsureFresh("PowerInfo", () => _mainCurrent, QueryPowerInfo);

        private float _mainCurrent2;
        public float MainCurrent2 => EnsureFresh("PowerInfo", () => _mainCurrent2, QueryPowerInfo);

        private float _mainCurrent3;
        public float MainCurrent3 => EnsureFresh("PowerInfo", () => _mainCurrent3, QueryPowerInfo);

        private float _outCurrent;
        public float OutCurrent => EnsureFresh("PowerInfo", () => _outCurrent, QueryPowerInfo);

        private float _bypassPF;
        public float BypassPF => EnsureFresh("PowerInfo", () => _bypassPF, QueryPowerInfo);

        private float _mainPF;
        public float MainPF => EnsureFresh("PowerInfo", () => _mainPF, QueryPowerInfo);

        private float _mainPF2;
        public float MainPF2 => EnsureFresh("PowerInfo", () => _mainPF2, QueryPowerInfo);

        private float _mainPF3;
        public float MainPF3 => EnsureFresh("PowerInfo", () => _mainPF3, QueryPowerInfo);

        private float _outputPF;
        public float OutputPF => EnsureFresh("PowerInfo", () => _outputPF, QueryPowerInfo);

        private SwitchState _currentSwitchState;
        public SwitchState CurrentSwitchState => EnsureFresh("SwitchesStatus", () => _currentSwitchState, QuerySwitchesStatus);

        private OutputStatus _currentOutputStatus;
        public OutputStatus CurrentOutputStatus => EnsureFresh("OutputStatus", () => _currentOutputStatus, QueryOutputStatus);

        private HardwareType _currentHardware;
        public HardwareType CurrentHardware => EnsureFresh("AutoDetect", () => _currentHardware, QueryAutoDetect, runOnce: true);

        private bool _batteryConnected;
        public bool BatteryConnected => EnsureFresh("BattCStatus", () => _batteryConnected, QueryBattCStatus);

        private SystemStatusFlags _statusFlags = SystemStatusFlags.None;
        public SystemStatusFlags StatusFlags => EnsureFresh("QueryStatus", () => _statusFlags, QueryStatus);

        private SystemStatus2Flags _statusFlags2 = SystemStatus2Flags.None;
        public SystemStatus2Flags StatusFlags2 => EnsureFresh("QueryStatus", () => _statusFlags2, QueryStatus);

        private SystemStatus3Flags _statusFlags3 = SystemStatus3Flags.None;
        public SystemStatus3Flags StatusFlags3 => EnsureFresh("QueryIPSCRStatus", () => _statusFlags3, QueryIPSCRStatus);

        public bool IsAmbientOverTemp => StatusFlags.HasFlag(SystemStatusFlags.AmbientOverTemp);
        public bool IsRecCanFail => StatusFlags.HasFlag(SystemStatusFlags.RecCanFail);
        public bool IsInvIoCanFail => StatusFlags.HasFlag(SystemStatusFlags.InvIoCanFail);
        public bool IsInvDataCanFail => StatusFlags.HasFlag(SystemStatusFlags.InvDataCanFail);
        public bool IsBypassPowerFuseFail => StatusFlags2.HasFlag(SystemStatus2Flags.BypassPowerFuseFail);
        public bool IsRatedKvaOverRange => StatusFlags2.HasFlag(SystemStatus2Flags.RatedKvaOverRange);

        public int BufferUnderRun = 0;
        public string UnderRunMethod = "";


        public RackPowerUpsClient(string portName, int baudRate)
        {
            _serialPort = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One)
            {
                ReadTimeout = 2000,
                WriteTimeout = 2000
            };
        }

        public void Open() => _serialPort.Open();
        public void Close() => _serialPort?.Close();
        public void Dispose()
        {
            Close();
            _serialPort?.Dispose();
            GC.SuppressFinalize(this);
        }

        private T EnsureFresh<T>(string groupKey, Func<T> getter, Action refresher, bool runOnce = false)
        {
            lock (_sync)
            {
                if (runOnce)
                {
                    // Only run refresher if value is null/default
                    if (!_lastUpdated.ContainsKey(groupKey))
                    {
                        refresher();
                        _lastUpdated[groupKey] = DateTime.MaxValue; // mark as "never refresh"
                    }
                    return getter();
                }

                DateTime last;
                if (!_lastUpdated.TryGetValue(groupKey, out last) || (DateTime.Now - last) > _maxAge)
                {
                    refresher(); // run query once for group
                    _lastUpdated[groupKey] = DateTime.Now;
                }
                return getter();
            }
        }

        public void QueryAutoDetect()
        {
            lock (_sync)
            {
                SendData("01044e74000166f8");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(50, 100, "QueryAutoDetect")));
                _currentHardware = (HardwareType)_regs[0];
                //_regs[0] 
                //1 ACM (10-600kVA)
                //2 TT (10-40kVA)
                //3 TS (10-20kVA)
                //4 SS (6-20kVA)
                //5 SS (1-3kVA)
            }
        }

        public void QueryManufacturerInfo()
        {
            lock (_sync)
            {
                SendData("010327110065df50");
                var frame = ModbusHelper.ParseFrame(ReadDataAdaptive(260, 100, "QueryManufacturerInfo"));
                _model = ModbusHelper.ByteToString(frame.Data, 143, 8);
                _regs = ModbusHelper.ExtractRegisters(frame);
                _batteryCount = _regs[10];
            }
        }

        public void QueryPowerInfo()
        {
            lock (_sync)
            {
                // Power Info
                //SendData("01044e210067f6c2");
                SendData("01044E210071770C"); //Read More registers
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(290, 100, "QueryPowerInfo")));
                if (_regs == null || _regs.Length < 112) // adjust min length
                {
                    throw new InvalidOperationException(
                        $"QueryPowerInfo Unexpected register length: got {_regs?.Length ?? 0}");
                }
                _bypassVoltage = _regs[0] * 0.1f;
                _bypassCurrent = _regs[3] * 0.1f;
                _bypassFreq = _regs[6] * 0.01f;
                _bypassPF = _regs[9] * 0.01f;
                _mainVoltage = _regs[12] * 0.1f;
                _mainVoltage2 = _regs[13] * 0.1f;
                _mainVoltage3 = _regs[14] * 0.1f;
                _mainCurrent = _regs[15] * 0.1f;
                _mainCurrent2 = _regs[16] * 0.1f;
                _mainCurrent3 = _regs[17] * 0.1f;
                _mainFreq = _regs[18] * 0.01f;
                _mainFreq2 = _regs[19] * 0.01f;
                _mainFreq3 = _regs[20] * 0.01f;
                _mainPF = _regs[21] * 0.01f;
                _mainPF2 = _regs[22] * 0.1f;
                _mainPF3 = _regs[23] * 0.1f;
                _outputVoltage = _regs[24] * 0.1f;
                _outCurrent = _regs[27] * 0.1f;
                _outputFreq = _regs[30] * 0.01f;
                _outputPF = _regs[33] * 0.01f;
                _va = _regs[36];
                _watts = _regs[39];
                _outputReactivePower = _regs[42] * 0.1f;
                _loadPercent = _regs[45] * 0.1f;
                _ambientTemp = _regs[48] * 0.1f;
                _batteryVoltage = _regs[49] * 0.1f;
                _batteryCurrent = _regs[51] * 0.1f;
                _batteryTemp = _regs[53] * 0.1f;
                _batteryTimeRemain = _regs[54] * 0.1f;
                _batteryCapacity = _regs[55] * 0.1f;
                _bypassFanHour = _regs[67];
                _dustFilterDays = _regs[68];
                _runtime = _regs[69] * 0.1f;
                _dischargehours = _regs[70] * 0.1f;
                _dischargeamount = _regs[72];
                _lockremaintime = _regs[74];
                _upskey = _regs[75];
                _currentHardware = (HardwareType)_regs[83];
                _upstype = _regs[84];
                _currentcorrectablemark = _regs[86];
                _inverteravgvolt = _regs[87] * 0.1f;
                _bypassavgvolt = _regs[90] * 0.1f;
                _syscode = _regs[93];
                _selfagingmode = _regs[94];
                _ratedInputVoltage = _regs[95];
                _ratedInputFreq = _regs[96];
                _ratedOutputVoltage = _regs[97];
                _ratedOutputFreq = _regs[98];
                _maxWatt = _regs[100];
                _busVoltage = _regs[102] * 0.1f;
                _chargerVoltage = _regs[104] * 0.1f;
                _chargerCurrent = _regs[106] * 0.1f;
                _inverterVoltage = _regs[108] * 0.1f;
                _fanruntime = _regs[111];
                _capacitortime = _regs[112];
            }
        }

        public void QueryVersionTemps()
        {
            lock (_sync)
            {
                // Version + Temps
                SendData("01044ea3000b56c7");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(70, 50, "QueryVersionTemps")));
                if (_regs == null || _regs.Length < 4) // adjust min length
                {
                    throw new InvalidOperationException(
                        $"QueryTemps Unexpected register length: got {_regs?.Length ?? 0}");
                }
                _version = _regs[7].ToString() + "." + _regs[8] + "." + _regs[9];
                _recIgbtTemp = _regs[0] * 0.1f;
                _invIgbtTemp = _regs[3] * 0.1f;
            }
        }

        // === Control Commands ===
        public void ClearFaults()
        {
            lock (_sync)
            {
                SendData("016828A100FF7801");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ClearFaults")));
            }
        }
        public void ManualBypass()
        {
            lock (_sync)
            {
                SendData("016828A400FF6800");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualBypass")));
            }
        }
        public void ManualTransferToInverter()
        {
            lock (_sync)
            {
                SendData("016828A40001E980");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualTransferToInverter")));
            }
        }
        public void ECSManualBypass()
        {
            lock (_sync)
            {
                SendData("016828A40002A981");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ECSManualBypass")));
            }
        }

        public void BatteryTest()
        {
            lock (_sync)
            {
                SendData("016828A5000F3984");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "BatteryTest")));
            }
        }
        public void BatteryMaintenance()
        {
            lock (_sync)
            {
                SendData("016828A500F079C4"); _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "BatteryMaintenance")));
            }
        }
        public void ManualFloat()
        {
            lock (_sync)
            {
                SendData("016828A5F0003D80"); _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualFloat")));
            }
        }
        public void ManualBoost()
        {
            lock (_sync)
            {
                SendData("016828A50F007C70"); _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualBoost")));
            }
        }

        public void StopTest()
        {
            lock (_sync)
            {
                SendData("016828A5FFFF7830");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "StopTest")));
            }
        }

        public void AlarmHistory()
        {
            lock (_sync)
            {
                SendData("016600010006D800");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "AlarmHistory")));
            }
        }

        public void QueryOutputStatus()
        {
            lock (_sync)
            {
                SendData("010400010001600a");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryOutputStatus")));
                _currentOutputStatus = (OutputStatus)_regs[0];
                //_regs[0] 0 No Out, 1 Inverter, 2 Dieode
            }
        }

        public void QueryBattCStatus()
        {
            lock (_sync)
            {
                SendData("010400020001900a");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryBattCStatus")));
                _batteryConnected = _regs[0] == 1;
                //_regs[0] 0 No Connection, 1 Connected
            }
        }

        public void QuerySwitchesStatus()
        {
            lock (_sync)
            {
                SendData("010400030001c1ca");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QuerySwitchesStatus")));
                _currentSwitchState = (SwitchState)_regs[0];
                //_regs[0] 
                //0 Bat Not Connected
                //1 Normal
                //2 MBCB Closed Bat Not Connected
                //3 MBCB Closed
                //4 EPO Batter Not Connected
                //5 EPO
                //6 MBCB Closed Bat Not Connected EPO
                //7 MBCB Closed EPO
            }
        }

        public void QueryStatus()
        {
            lock (_sync)
            {
                SendData("01040004000631c9");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(70, 100, "QueryStatus")));
                int val4 = (_regs != null && _regs.Length > 4) ? _regs[4] : 0;
                int val5 = (_regs != null && _regs.Length > 5) ? _regs[5] : 0;
                _statusFlags = (SystemStatusFlags)val4;
                _statusFlags2 = (SystemStatus2Flags)val5;
                //_regs[4] 
                //0 Normal
                //1 Ambient Over Temp
                //2 REC CAN Fail
                //3 1 and 2
                //4 INV IO CAN Fail
                //5 1 and 4
                //6 2 and 4
                //7 1 and 2 and 4
                //8 INV DATA CAN Fail
                //9 1 and 8
                //A 2 and 8
                //B 1 and 2 and 8
                //C 4 and 8
                //D 1 and 4 and 8
                //E 2 and 4 and 8
                //_regs[5] 
                //0 Normal
                //1 Byp Power Fuse Fail
                //2 Rated KVA Over Range
                //3 1 and 2
            }

        }

        public void QueryIPSCRStatus()
        {
            lock (_sync)
            {
                SendData("010400c900036035");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryIPSCRStatus")));
                int val1 = (_regs != null && _regs.Length > 2) ? _regs[2] : 0;
                _statusFlags3 = (SystemStatus3Flags)val1;
                //_regs[2] 
                //0 Normal
                //2 No IP SCR TmpSensor
                //4 IP SCR Over Temp
                //6 2 and 4
            }
        }

        public void QueryStatus3()
        {
            lock (_sync)
            {
                SendData("0104000a000cd00d");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryStatus3")));
            }
        }

        public void QueryStatus5()
        {
            lock (_sync)
            {
                SendData("010400020001900a");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryStatus5")));
            }
        }

        public void QueryStatus6()
        {
            lock (_sync)
            {
                SendData("01044ee90002b717");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryStatus6")));
            }
        }

        public void QueryStatus7()
        {
            lock (_sync)
            {
                SendData("01040000000131ca");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryStatus7")));
            }
        }

        public void QueryStatus8()
        {
            lock (_sync)
            {
                SendData("010327760004af67");
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryStatus8")));
            }
        }

        public void SetBacklightTimer(int minutes)
        {
            lock (_sync)
            {
                switch (minutes)
                {
                    case 1: SendData("016827400001AAA3"); break;
                    case 3: SendData("0168274000032B62"); break;
                    case 5: SendData("016827400005AB60"); break;
                    case 10: SendData("01682740000AEB64"); break;
                    case 20: SendData("0168274000146B6C"); break;
                    case 30: SendData("01682740001EEB6B"); break;
                    default: throw new ArgumentOutOfRangeException(nameof(minutes), "Valid values: 1,3,5,10,20,30");
                }
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(500, 100, "SetBacklightTimer")));
            }
        }

        // === Internal Helpers ===
        public void SendData(string hexString)
        {
            if (_serialPort == null || !_serialPort.IsOpen)
                throw new InvalidOperationException("Serial port is not open");

            byte[] buffer = Enumerable.Range(0, hexString.Length / 2).Select(x => Convert.ToByte(hexString.Substring(x * 2, 2), 16)).ToArray();
            _serialPort.Write(buffer, 0, buffer.Length);
        }

        public byte[] ReadDataAdaptive(int baseDelayMs, int maxExtraMs, string method)
        {
            // Start with base delay
            Thread.Sleep(baseDelayMs);
            if (_serialPort.IsOpen)
            {
                int waited = 0;
                int lastCount = -1;

                // Keep waiting in small steps until data stops increasing or cap reached
                while (waited < maxExtraMs)
                {
                    int count = _serialPort.BytesToRead;
                    if (count > 0 && count == lastCount)
                        break; // no more data is coming

                    lastCount = count;
                    Thread.Sleep(10); // small step
                    waited += 10;
                }

                if (waited > 10)
                {
                    BufferUnderRun++;
                    UnderRunMethod = method;
                }

                int bytesToRead = _serialPort.BytesToRead;
                if (bytesToRead == 0)
                    throw new TimeoutException("No response from UPS");

                if (bytesToRead > 0)
                {
                    byte[] response = new byte[bytesToRead];
                    _serialPort.Read(response, 0, response.Length);
                    return response;
                }
            }
            return null;
        }
    }
}