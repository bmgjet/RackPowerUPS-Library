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
        public ModbusFrame RequestFrame(byte slaveAddress, byte functionCode, ushort startRegister, ushort amountToRead)
        {
            SlaveAddress = slaveAddress;
            FunctionCode = functionCode;
            Data = new byte[4];
            Data[0] = (byte)(startRegister >> 8);
            Data[1] = (byte)(startRegister & 0xFF);
            Data[2] = (byte)(amountToRead >> 8);
            Data[3] = (byte)(amountToRead & 0xFF);
            byte[] frameWithoutCrc = new byte[6];
            frameWithoutCrc[0] = SlaveAddress;
            frameWithoutCrc[1] = FunctionCode;
            Array.Copy(Data, 0, frameWithoutCrc, 2, Data.Length);
            ushort crc = ModbusHelper.Crc16(frameWithoutCrc);
            Crc = crc;
            return this;
        }

        public byte SlaveAddress { get; set; }
        public byte FunctionCode { get; set; }
        public byte[] Data { get; set; } = new byte[0];
        public ushort Crc { get; set; }
        public bool CrcValid { get; set; }

        public byte[] ToBytes()
        {
            byte[] frame = new byte[Data.Length + 4];
            frame[0] = SlaveAddress;
            frame[1] = FunctionCode;
            Array.Copy(Data, 0, frame, 2, Data.Length);
            frame[frame.Length - 2] = (byte)(Crc & 0xFF);
            frame[frame.Length - 1] = (byte)(Crc >> 8);
            return frame;
        }
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
        public byte SlaveID;
        public ushort[] _regs;

        private readonly object _sync = new object();
        private readonly Dictionary<string, DateTime> _lastUpdated = new Dictionary<string, DateTime>();
        public TimeSpan _maxAge = TimeSpan.FromMilliseconds(300);
        private readonly Dictionary<string, ushort> _reverseMap = _registerMap.ToDictionary(kv => kv.Value, kv => kv.Key, StringComparer.OrdinalIgnoreCase);
        public static Dictionary<ushort, string> _registerMap = new Dictionary<ushort, string>
        {
    { 20001, "AC bypass voltage ph_A" },
    { 20004, "AC bypass current ph_A" },
    { 20007, "AC bypass frequency ph_A" },
    { 20010, "AC Bypass PF_A" },
    { 20013, "AC input voltage ph_A" },
    { 20014, "AC input voltage ph_B" },
    { 20015, "AC input voltage ph_C" },
    { 20016, "AC input current ph_A" },
    { 20017, "AC input current ph_B" },
    { 20018, "AC input current ph_C" },
    { 20019, "AC input frequency ph_A" },
    { 20020, "AC input frequency ph_B" },
    { 20021, "AC input frequency ph_C" },
    { 20022, "AC Input PF_A" },
    { 20023, "AC Input PF_B" },
    { 20024, "AC Input PF_C" },
    { 20025, "AC output voltage ph_A" },
    { 20028, "AC output current ph_A" },
    { 20031, "AC output frequency ph_A" },
    { 20034, "AC output PF_A" },
    { 20037, "Output apparent power ph_A" },
    { 20040, "Output active power ph_A" },
    { 20043, "Output reactive power ph_A" },
    { 20046, "Load percentage ph_A" },
    { 20049, "Ambient temperature" },
    { 20050, "Positive battery pack voltage" },
    { 20052, "Positive battery pack current" },
    { 20054, "Battery temperature" },
    { 20055, "Battery time remaining" },
    { 20056, "Battery capacity" },
    { 20068, "Bypass fan running time" },
    { 20069, "Dust filter used days" },
    { 20070, "Total battery runtime" },
    { 20071, "Total battery discharge time" },
    { 20073, "Battery discharge times" },
    { 20075, "UPS lock remaining time" },
    { 20076, "UPS Key" },
    { 20084, "UPS model" },
    { 20085, "UPS type" },
    { 20087, "Current correctable mark" },
    { 20088, "Average module inverter voltage A" },
    { 20091, "Average module bypass voltage A" },
    { 20094, "System code" },
    { 20096, "Rated input voltage" },
    { 20097, "Rated input frequency" },
    { 20098, "Rated output voltage" },
    { 20099, "Rated output frequency" },
    { 20101, "Module rated power" },
    { 20103, "DC bus voltage +" },
    { 20104, "DC bus voltage -" },
    { 20105, "Charger voltage +" },
    { 20106, "Charger voltage -" },
    { 20107, "Charging current +" },
    { 20108, "Charging current -" },
    { 20109, "Inverter voltage ph_A" },
    { 20112, "Fan running time" },
    { 20113, "Bus capacitor operating time" },
    { 20114, "Rectified input IO" },
    { 20115, "Rectified output IO" },
    { 20118, "Rectifier prohibited start-up sign summary" },
    { 20119, "Comprehensive list of mains and battery symbols" },
    { 20123, "Inverter shutdown lockout sign summary" },
    { 20131, "Rectification IGBT temperature" },
    { 20134, "Inverter IGBT temperature" },
    { 20138, "Rectification version number" }, 
    { 20142, "Rectification enters SERVICE mode flag" },
    { 20143, "Input current ADC sampling value A+" },
    { 20144, "Input current ADC sampling value B+" },
    { 20145, "Input current ADC sampling value C+" },
    { 20146, "Input current ADC sampling value A-" },
    { 20147, "Input current ADC sampling value B-" },
    { 20148, "Input current ADC sampling value C-" },
    { 20149, "Input voltage ADC sampling value A" },
    { 20150, "Input voltage ADC sampling value B" },
    { 20151, "Input voltage ADC sampling value C" },
    { 20152, "Bus voltage ADC sampling value +" },
    { 20153, "Bus voltage ADC sampling value -" },
    { 20154, "Charger current ADC sampling value +" },
    { 20156, "Charger voltage ADC sampling value +" },
    { 20158, "Battery voltage ADC sampling value +" },
    { 20161, "Battery temperature ADC sampling value" },
    { 20167, "Rectification temperature" },
    { 20168, "Inversion temperature" },
    { 20201, "History storage pointer" },
    { 20202, "Number of historical records stored" },
    { 20207, "Battery deep discharge times" },
    { 20208, "Battery shallow discharge times" },
    { 20210, "Battery SOH" }
};



        public enum PowerSupplyMode
        {
            NoPowerSupply = 0,
            UPSPowered = 1,
            BypassPowerSupply = 2
        }

        public enum BatteryStatus
        {
            None = 0,
            Connected = 1,
            FloatCharge = 2,
            EqualizeCharge = 3
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
        public enum StateStatusFlags
        {
            SoftStart =  0,
            Normal = 1 << 0,
            None = 1 << 1
        }

        [Flags]
        public enum AlarmStatusFlags
        {
            None = 0,
            Warning = 1 << 0
        }

        [Flags]
        public enum BypassStatusFlags
        {
            None = 1 << 0,
            BypassPowerSupply = 1 << 1,
            Normal = 1 << 2
        }

        [Flags]
        public enum BatterySelfTestFlags
        {
            None = 0,
            Success = 1 << 0,
            Fail = 1 << 1,
            MaintenanceTest = 1 << 2
        }

        [Flags]
        public enum BatteryMaintenanceFlags
        {
            None = 0,
            Success = 1 << 0,
            Fail = 1 << 1,
            MaintenanceTest = 1 << 2
        }

        [Flags]
        public enum SystemStatusFlags
        {
            Normal = 0,
            AmbientOverTemp = 1 << 0,
            RecCanFail = 1 << 1,
            InvIOCanFail = 1 << 2,
            InvDataCanFail = 1 << 3
        }

        [Flags]
        public enum SystemStatus2Flags
        {
            Normal = 0,
            BypassFuseFail = 1 << 0,
            RatedKVAOverRange = 1 << 1
        }

        [Flags]
        public enum SystemStatus3Flags
        {
            Normal = 0,
            NoIpScrTempSensor = 1 << 1,
            IpScrOverTemp = 1 << 2,
            LowBatteryLife = 1 << 7
        }

        [Flags]
        public enum SystemStatus4Flags
        {
            Normal = 0,
            RectificationFault = 1 << 0,
            InverterFault = 1 << 1,
            RectifierOverTemp = 1 << 2,
            FanFailure = 1 << 3,
            OutputOverload = 1 << 4,
            OverloadTimeout = 1 << 5,
            InverterOverTemp = 1 << 6,
            InverterProtection = 1 << 7,
            ManualShutdown = 1 << 8,
            BatteryChargerFailure = 1 << 9,
            CurrentSharingAbnormality = 1 << 10,
            SynchronousSignalAbnormality = 1 << 11,
            InputVoltageAbnormality = 1 << 12,
            BatteryVoltageAbnormality = 1 << 13,
            OutputVoltageAbnormality = 1 << 14,
            BypassVoltageAbnormality = 1 << 15,
        }

        [Flags]
        public enum SystemStatus5Flags
        {
            Normal = 0,
            UnbalancedInputCurrent = 1 << 1,
            BusbarOvervoltage = 1 << 2,
            RectificationSoftStartFailed = 1 << 3,
            InverterRelayOpenCircuit = 1 << 4,
            InverterSwitchShortCircuit = 1 << 5,
            PWMTrackingSignalAbnormality = 1 << 6
        }

        [Flags]
        public enum CabinetAlarmStatus_flags
        {
            BatteryConnectionStatus = 0,
            None = 1,
            MaintenanceBypassCircuitBreakerStatus = 2,
            Epo = 4,
            InsufficientInverterStartupCapacity = 8,
            GeneratorAccess = 16,
            AcPowerAbnormality = 32,
            BypassPhaseSequenceReverse = 64,
            BypassVoltageAbnormality = 128,
            BypassFault = 256,
            BypassOverload = 512,
            BypassOverloadTimeout = 1024,
            BypassSuperTracking = 2048,
            SwitchingTimesReached = 4096,
            OutputShortCircuit = 8192,
            BatteryEod = 16384,
            BatteryTestStartReserved = 32768,
        }

        [Flags]
        public enum CabinetAlarmStatus2_flags
        {
            None = 0,
            StopBatteryTest = 1,
            TroubleShooting = 2,
            HistroyClear = 4,
            ProhibitPowerOn = 8,
            ManualBypass = 16,
            BatteryLowVoltage = 32,
            BatteryConnectedReversely = 64,
            InputNLineDisconnected = 128,
            Reserved = 256,
            Reserved1 = 512,
            Reserved2 = 1024,
            ScheduleShutdown = 2048,
            Reserved3 = 4096,
            EODSystemProhibited = 8192,
            ParrallelLineFault = 16384,
            SignalLineFailure = 32768,
        }

        [Flags]
        public enum CabinetAlarmStatus3_flags
        {
            None = 0,
            Reserved = 1,
            Reserved1 = 2,
            Reserved2 = 4,
            Reserved3 = 8,
            Reserved4 = 16,
            Reserved5 = 32,
            Reserved6 = 64,
            Reserved7 = 128,
            Reserved8 = 256,
            Reserved9 = 512,
            Reserved10 = 1024,
            Reserved11 = 2048,
            Reserved12 = 4096,
            Reserved13 = 8192,
            Reserved14 = 16384,
            Reserved15 = 32768,
        }

        [Flags]
        public enum CabinetAlarmStatus4_flags
        {
            None = 0,
            Reserved = 1,
            Reserved1 = 2,
            Reserved2 = 4,
            Reserved3 = 8,
            Reserved4 = 16,
            Reserved5 = 32,
            Reserved6 = 64,
            Reserved7 = 128,
            Reserved8 = 256,
            Reserved9 = 512,
            Reserved10 = 1024,
            Reserved11 = 2048,
            Reserved12 = 4096,
            BMSCommunicationFailure = 8192,
            Reserved14 = 16384,
            Reserved15 = 32768,
        }

        public bool IsSoftStart => EnsureFresh("QueryIPSCRStatus", () => _statusFlags5.HasFlag(SystemStatus5Flags.RectificationSoftStartFailed), QueryIPSCRStatus);
        public bool IsNormalState => EnsureFresh("QueryIPSCRStatus", () => _statusFlags5.HasFlag(SystemStatus5Flags.Normal), QueryIPSCRStatus);
        public bool IsWarning => EnsureFresh("QueryStatusFlags", () => AlarmFlags.HasFlag(AlarmStatusFlags.Warning), QueryStatusFlags);
        public bool IsBypassPowerSupply => EnsureFresh("QueryStatusFlags", () => _BypassFlags.HasFlag(BypassStatusFlags.BypassPowerSupply), QueryStatusFlags);
        public bool IsNormalBypass => EnsureFresh("QueryStatusFlags", () => _BypassFlags.HasFlag(BypassStatusFlags.Normal), QueryStatusFlags);
        public bool IsBatterySelfTestSuccess => EnsureFresh("QueryStatusFlags", () => BatterySelfCheck_Flags.HasFlag(BatterySelfTestFlags.Success), QueryStatusFlags);
        public bool IsBatterySelfTestFail => EnsureFresh("QueryStatusFlags", () => BatterySelfCheck_Flags.HasFlag(BatterySelfTestFlags.Fail), QueryStatusFlags);
        public bool IsBatterySelfTestMaintenance => EnsureFresh("QueryStatusFlags", () => BatterySelfCheck_Flags.HasFlag(BatterySelfTestFlags.MaintenanceTest), QueryStatusFlags);
        public bool IsBatteryMaintenanceSuccess => EnsureFresh("QueryStatusFlags2", () => BatteryMaintenanceCheck_Flags.HasFlag(BatteryMaintenanceFlags.Success), QueryStatusFlags2);
        public bool IsBatteryMaintenanceFail => EnsureFresh("QueryStatusFlags2", () => BatteryMaintenanceCheck_Flags.HasFlag(BatteryMaintenanceFlags.Fail), QueryStatusFlags2);
        public bool IsBatteryMaintenanceTest => EnsureFresh("QueryStatusFlags2", () => BatteryMaintenanceCheck_Flags.HasFlag(BatteryMaintenanceFlags.MaintenanceTest), QueryStatusFlags2);
        public bool IsAmbientOverTemp => EnsureFresh("QueryStatusFlags2", () => StatusFlags.HasFlag(SystemStatusFlags.AmbientOverTemp), QueryStatusFlags2);
        public bool IsRecCanFail => EnsureFresh("QueryStatusFlags2", () => StatusFlags.HasFlag(SystemStatusFlags.RecCanFail), QueryStatusFlags2);
        public bool IsInvIOCanFail => EnsureFresh("QueryStatusFlags2", () => StatusFlags.HasFlag(SystemStatusFlags.InvIOCanFail), QueryStatusFlags2);
        public bool IsInvDataCanFail => EnsureFresh("QueryStatusFlags2", () => StatusFlags.HasFlag(SystemStatusFlags.InvDataCanFail), QueryStatusFlags2);
        public bool IsBypassFuseFail => EnsureFresh("QueryStatusFlags2", () => StatusFlags2.HasFlag(SystemStatus2Flags.BypassFuseFail), QueryStatusFlags2);
        public bool IsRatedKVAOverRange => EnsureFresh("QueryStatusFlags2", () => StatusFlags2.HasFlag(SystemStatus2Flags.RatedKVAOverRange), QueryStatusFlags2);
        public bool IsNoIpScrTempSensor => EnsureFresh("QueryIPSCRStatus", () => StatusFlags3.HasFlag(SystemStatus3Flags.NoIpScrTempSensor), QueryIPSCRStatus);
        public bool IsIpScrOverTemp => EnsureFresh("QueryIPSCRStatus", () => StatusFlags3.HasFlag(SystemStatus3Flags.IpScrOverTemp), QueryIPSCRStatus);
        public bool IsLowBatteryLife => EnsureFresh("QueryIPSCRStatus", () => StatusFlags3.HasFlag(SystemStatus3Flags.LowBatteryLife), QueryIPSCRStatus);
        public bool IsInverterFault => EnsureFresh("QueryIPSCRStatus", () => StatusFlags4.HasFlag(SystemStatus4Flags.InverterFault), QueryIPSCRStatus);
        public bool IsRectifierOverTemp => EnsureFresh("QueryIPSCRStatus", () => StatusFlags4.HasFlag(SystemStatus4Flags.RectifierOverTemp), QueryIPSCRStatus);
        public bool IsFanFailure => EnsureFresh("QueryIPSCRStatus", () => StatusFlags4.HasFlag(SystemStatus4Flags.FanFailure), QueryIPSCRStatus);
        public bool IsInvOverTemp => EnsureFresh("QueryIPSCRStatus", () => StatusFlags4.HasFlag(SystemStatus4Flags.InverterOverTemp), QueryIPSCRStatus);
        public bool IsInvOverLoad => EnsureFresh("QueryIPSCRStatus", () => StatusFlags4.HasFlag(SystemStatus4Flags.OutputOverload), QueryIPSCRStatus);
        public bool IsBusbarOvervoltage => EnsureFresh("QueryIPSCRStatus", () => StatusFlags5.HasFlag(SystemStatus5Flags.BusbarOvervoltage), QueryIPSCRStatus);
        public bool IsInverterRelayOpenCircuit => EnsureFresh("QueryIPSCRStatus", () => StatusFlags5.HasFlag(SystemStatus5Flags.InverterRelayOpenCircuit), QueryIPSCRStatus);
        public bool IsInverterSwitchShortCircuit => EnsureFresh("QueryIPSCRStatus", () => StatusFlags5.HasFlag(SystemStatus5Flags.InverterSwitchShortCircuit), QueryIPSCRStatus);
        public bool IsPWMTrackingSignalAbnormality => EnsureFresh("QueryIPSCRStatus", () => StatusFlags5.HasFlag(SystemStatus5Flags.PWMTrackingSignalAbnormality), QueryIPSCRStatus);
        public bool IsRectificationSoftStartFailed => EnsureFresh("QueryIPSCRStatus", () => StatusFlags5.HasFlag(SystemStatus5Flags.RectificationSoftStartFailed), QueryIPSCRStatus);
        public bool IsUnbalancedInputCurrent => EnsureFresh("QueryIPSCRStatus", () => StatusFlags5.HasFlag(SystemStatus5Flags.UnbalancedInputCurrent), QueryIPSCRStatus);
        private string _model;
        public string Model => EnsureFresh("ManufacturerInfo", () => _model, QueryManufacturerInfo, runOnce: true);
        private string _version;
        public string Version => EnsureFresh("VersionTemps", () => _version, QueryVersionTemps, runOnce: true);
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
        public float RatedOutputVoltage => EnsureFresh("PowerInfo2", () => _ratedOutputVoltage, QueryPowerInfo, runOnce: true);
        private float _ratedOutputFreq;
        public float RatedOutputFreq => EnsureFresh("PowerInfo2", () => _ratedOutputFreq, QueryPowerInfo, runOnce: true);
        private float _ratedInputVoltage;
        public float RatedInputVoltage => EnsureFresh("PowerInfo2", () => _ratedInputVoltage, QueryPowerInfo, runOnce: true);
        private float _ratedInputFreq;
        public float RatedInputFreq => EnsureFresh("PowerInfo2", () => _ratedInputFreq, QueryPowerInfo, runOnce: true);
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
        public float RecIGBTTemp => EnsureFresh("Temps", () => _recIgbtTemp, QueryVersionTemps);
        private float _invIgbtTemp;
        public float InvIGBTTemp => EnsureFresh("Temps", () => _invIgbtTemp, QueryVersionTemps);
        private float _maxWatt;
        public float MaxWatt => EnsureFresh("PowerInfo2", () => _maxWatt, QueryPowerInfo, runOnce: true);
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
        private PowerSupplyMode _currentPowerSupplyStatus;
        public PowerSupplyMode CurrentPowerSupplyStatus => EnsureFresh("QueryStatusFlags", () => _currentPowerSupplyStatus, QueryStatusFlags);
        private HardwareType _currentHardware;
        public HardwareType CurrentHardware => EnsureFresh("AutoDetect", () => _currentHardware, QueryAutoDetect, runOnce: true);
        private BatteryStatus _batteryConnected;
        public BatteryStatus BatteryConnected => EnsureFresh("QueryStatusFlags", () => _batteryConnected, QueryStatusFlags);
        private CabinetAlarmStatus_flags _cabinetAlarmStatus_flags = CabinetAlarmStatus_flags.None;
        public CabinetAlarmStatus_flags CabinetAlarmStatus => EnsureFresh("QueryStatusFlags", () => _cabinetAlarmStatus_flags, QueryStatusFlags);
        private CabinetAlarmStatus2_flags _cabinetAlarmStatus2_flags = CabinetAlarmStatus2_flags.None;
        public CabinetAlarmStatus2_flags CabinetAlarmStatus2 => EnsureFresh("QueryStatusFlags2", () => _cabinetAlarmStatus2_flags, QueryStatusFlags2);
        private CabinetAlarmStatus3_flags _cabinetAlarmStatus3_flags = CabinetAlarmStatus3_flags.None;
        public CabinetAlarmStatus3_flags CabinetAlarmStatus3 => EnsureFresh("QueryStatusFlags2", () => _cabinetAlarmStatus3_flags, QueryStatusFlags2);
        private CabinetAlarmStatus4_flags _cabinetAlarmStatus4_flags = CabinetAlarmStatus4_flags.None;
        public CabinetAlarmStatus4_flags CabinetAlarmStatus4 => EnsureFresh("QueryStatusFlags2", () => _cabinetAlarmStatus4_flags, QueryStatusFlags2);
        private BatterySelfTestFlags _BatterySelfCheck_Flags = BatterySelfTestFlags.None;
        public BatterySelfTestFlags BatterySelfCheck_Flags => EnsureFresh("QueryStatusFlags", () => _BatterySelfCheck_Flags, QueryStatusFlags);
        private BatteryMaintenanceFlags _BatteryMaintenanceCheck_Flags = BatteryMaintenanceFlags.None;
        public BatteryMaintenanceFlags BatteryMaintenanceCheck_Flags => EnsureFresh("QueryStatusFlags2", () => _BatteryMaintenanceCheck_Flags, QueryStatusFlags2);
        private SystemStatusFlags _statusFlags = SystemStatusFlags.Normal;
        public SystemStatusFlags StatusFlags => EnsureFresh("QueryStatusFlags2", () => _statusFlags, QueryStatusFlags2);
        private SystemStatus2Flags _statusFlags2 = SystemStatus2Flags.Normal;
        public SystemStatus2Flags StatusFlags2 => EnsureFresh("QueryStatusFlags2", () => _statusFlags2, QueryStatusFlags2);
        private SystemStatus3Flags _statusFlags3 = SystemStatus3Flags.Normal;
        public SystemStatus3Flags StatusFlags3 => EnsureFresh("QueryIPSCRStatus", () => _statusFlags3, QueryIPSCRStatus);
        private SystemStatus4Flags _statusFlags4 = SystemStatus4Flags.Normal;
        public SystemStatus4Flags StatusFlags4 => EnsureFresh("QueryIPSCRStatus", () => _statusFlags4, QueryIPSCRStatus);
        private StateStatusFlags _RecstatusFlags = StateStatusFlags.Normal;
        public StateStatusFlags RecstatusFlags => EnsureFresh("QueryStatusFlags2", () => _RecstatusFlags, QueryStatusFlags2);
        private StateStatusFlags _InvstatusFlags = StateStatusFlags.Normal;
        public StateStatusFlags InvstatusFlags => EnsureFresh("QueryStatusFlags2", () => _InvstatusFlags, QueryStatusFlags2);
        private BypassStatusFlags _BypassFlags = BypassStatusFlags.Normal;
        public BypassStatusFlags BypassFlags => EnsureFresh("QueryStatusFlags2", () => _BypassFlags, QueryStatusFlags2);
        private AlarmStatusFlags _AlarmFlags = AlarmStatusFlags.None;
        public AlarmStatusFlags AlarmFlags => EnsureFresh("QueryAlarmStatus", () => _AlarmFlags, QueryAlarmStatus);

        private SystemStatus5Flags _statusFlags5 = SystemStatus5Flags.Normal;
        public SystemStatus5Flags StatusFlags5 => EnsureFresh("QueryIPSCRStatus", () => _statusFlags5, QueryIPSCRStatus);
        public int BufferUnderRun = 0;
        public string UnderRunMethod = "";

        public RackPowerUpsClient(string portName, int baudRate, byte slaveID = 01)
        {
            SlaveID = slaveID;
            _serialPort = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One)
            {
                ReadTimeout = 2000,
                WriteTimeout = 2000
            };
        }

        #region Helpers
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
                    if (!_lastUpdated.ContainsKey(groupKey))
                    {
                        refresher();
                        _lastUpdated[groupKey] = DateTime.MaxValue;
                    }
                    return getter();
                }

                DateTime last;
                if (!_lastUpdated.TryGetValue(groupKey, out last) || (DateTime.Now - last) > _maxAge)
                {
                    refresher();
                    _lastUpdated[groupKey] = DateTime.Now;
                }
                return getter();
            }
        }

        public void SendCommand(byte[] commandHeader, byte[] payload)
        {
            if (commandHeader == null || commandHeader.Length == 0) { throw new ArgumentException("SendCommand Command header cannot be null or empty."); }
            if (payload == null) { payload = new byte[0]; }
            byte[] frameWithoutCrc = new byte[1 + commandHeader.Length + payload.Length];
            frameWithoutCrc[0] = SlaveID;
            Array.Copy(commandHeader, 0, frameWithoutCrc, 1, commandHeader.Length);
            Array.Copy(payload, 0, frameWithoutCrc, 1 + commandHeader.Length, payload.Length);
            ushort crc = ModbusHelper.Crc16(frameWithoutCrc);
            byte[] frame = new byte[frameWithoutCrc.Length + 2];
            Array.Copy(frameWithoutCrc, frame, frameWithoutCrc.Length);
            frame[frame.Length - 2] = (byte)(crc & 0xFF);
            frame[frame.Length - 1] = (byte)(crc >> 8);
            _serialPort.Write(frame, 0, frame.Length);
        }

        public void SendRequest(byte FunctionCode, ushort Register, ushort NumberToRead = 1)
        {
            var frame = new ModbusFrame().RequestFrame(SlaveID, FunctionCode, Register, NumberToRead).ToBytes();
            if (frame == null) { throw new InvalidOperationException("SendRequest null frame."); }
            _serialPort.Write(frame, 0, frame.Length);
        }

        public string LookUpRegister(ushort registerID)
        {
            return _registerMap.TryGetValue(registerID, out var name) ? name : "Unknown Register";
        }

        public ushort LookUpRegister(string registerName)
        {
            if (_reverseMap.TryGetValue(registerName, out var id)) { return id; }
            var bestMatch = _reverseMap.Keys.OrderBy(k => FindClosestMatch(k, registerName)).FirstOrDefault();
            return bestMatch != null ? _reverseMap[bestMatch] : (ushort)0;
        }

        private static int FindClosestMatch(string s, string t)
        {
            if (string.IsNullOrEmpty(s)) return t.Length;
            if (string.IsNullOrEmpty(t)) return s.Length;
            var d = new int[s.Length + 1, t.Length + 1];
            for (int i = 0; i <= s.Length; i++) d[i, 0] = i;
            for (int j = 0; j <= t.Length; j++) d[0, j] = j;
            for (int i = 1; i <= s.Length; i++)
            {
                for (int j = 1; j <= t.Length; j++)
                {
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;
                    d[i, j] = new int[] {
                    d[i - 1, j] + 1,        
                    d[i, j - 1] + 1,        
                    d[i - 1, j - 1] + cost 
                }.Min();
                }
            }
            return d[s.Length, t.Length];
        }
        public byte[] ReadDataAdaptive(int baseDelayMs, int maxExtraMs, string method)
        {
            // Start with base delay
            Thread.Sleep(baseDelayMs);
            if (_serialPort.IsOpen)
            {
                int waited = 0;
                int lastCount = -1;
                while (waited < maxExtraMs)
                {
                    int count = _serialPort.BytesToRead;
                    if (count > 0 && count == lastCount) { break; }
                    lastCount = count;
                    Thread.Sleep(10);
                    waited += 10;
                }

                if (waited > 10)
                {
                    BufferUnderRun++;
                    UnderRunMethod = method;
                }

                int bytesToRead = _serialPort.BytesToRead;
                if (bytesToRead == 0) { throw new TimeoutException("No response from UPS"); }
                if (bytesToRead > 0)
                {
                    byte[] response = new byte[bytesToRead];
                    _serialPort.Read(response, 0, response.Length);
                    return response;
                }
            }
            return null;
        }
        #endregion


        #region Single Register Requests

        public void QueryAutoDetect() //Register 20084
        {
            lock (_sync)
            {
                SendRequest(04, 20084, 1);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(50, 100, "QueryAutoDetect")));
                if (_regs == null || _regs.Length < 1) { throw new InvalidOperationException($"QueryAutoDetect Unexpected register length: got {_regs?.Length ?? 0}"); }
                _currentHardware = (HardwareType)_regs[0];
            }
        }

        public void QueryManufacturerInfo() //Register 10001
        {
            lock (_sync)
            {
                SendRequest(03, 10001, 101);
                var frame = ModbusHelper.ParseFrame(ReadDataAdaptive(255, 100, "QueryManufacturerInfo"));
                _regs = ModbusHelper.ExtractRegisters(frame);
                if (_regs == null || _regs.Length < 11) { throw new InvalidOperationException($"QueryManufacturerInfo Unexpected register length: got {_regs?.Length ?? 0}"); }
                _model = ModbusHelper.ByteToString(frame.Data, 143, 8);
                _regs = ModbusHelper.ExtractRegisters(frame);
                _batteryCount = _regs[10];
            }
        }

        public void QueryAlarmStatus() //Register 22
        {
            lock (_sync)
            {
                SendRequest(04, 22, 1);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryAlarmStatus")));
                if (_regs == null || _regs.Length < 1) { throw new InvalidOperationException($"QueryAlarmStatus Unexpected register length: got {_regs?.Length ?? 0}"); }
                _AlarmFlags = (AlarmStatusFlags)_regs[0];
            }
        }

        public void AlarmHistory()
        {
            lock (_sync)
            {
                SendRequest(66, 1, 6);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "AlarmHistory")));
            }
        }
        #endregion

        #region Grouped Register Requests
        public void QueryPowerInfo() //Registers 20001 - 20113
        {
            lock (_sync)
            {
                SendRequest(04, 20001, 113);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(285, 100, "QueryPowerInfo")));
                if (_regs == null || _regs.Length < 113) { throw new InvalidOperationException($"QueryPowerInfo Unexpected register length: got {_regs?.Length ?? 0}"); }
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

        public void QueryVersionTemps() //Registers 20131 - 20141
        {
            lock (_sync)
            {
                SendRequest(04, 20131, 10);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(65, 100, "QueryVersionTemps")));
                if (_regs == null || _regs.Length < 10) { throw new InvalidOperationException($"QueryVersionTemps Unexpected register length: got {_regs?.Length ?? 0}"); }
                _recIgbtTemp = _regs[0] * 0.1f;
                _invIgbtTemp = _regs[3] * 0.1f;
                _version = _regs[7].ToString() + "." + _regs[8] + "." + _regs[9];
            }
        }

        public void QueryStatusFlags() //Register 1 - 4
        {
            lock (_sync)
            {
                SendRequest(04, 1, 4);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryStatusFlags")));
                if (_regs == null || _regs.Length < 4) { throw new InvalidOperationException($"QueryStatusFlags Unexpected register length: got {_regs?.Length ?? 0}"); }
                _currentPowerSupplyStatus = (PowerSupplyMode)_regs[0];
                _batteryConnected = (BatteryStatus)_regs[1];
                _cabinetAlarmStatus_flags = (CabinetAlarmStatus_flags)_regs[2];
                _BatterySelfCheck_Flags = (BatterySelfTestFlags)_regs[3];
            }
        }

        public void QueryStatusFlags2() //Register 6 - 13
        {
            lock (_sync)
            {
                SendRequest(04, 6, 8);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(65, 100, "QueryStatusFlags2")));
                if (_regs == null || _regs.Length < 8) { throw new InvalidOperationException($"QueryStatusFlags2 Unexpected register length: got {_regs?.Length ?? 0}"); }
                _BatteryMaintenanceCheck_Flags = (BatteryMaintenanceFlags)_regs[0];
                _cabinetAlarmStatus2_flags = (CabinetAlarmStatus2_flags)_regs[1];
                _cabinetAlarmStatus3_flags = (CabinetAlarmStatus3_flags)_regs[2];
                _cabinetAlarmStatus4_flags = (CabinetAlarmStatus4_flags)_regs[3];
                _RecstatusFlags = (StateStatusFlags)_regs[4];
                _InvstatusFlags = (StateStatusFlags)_regs[5];
                _BypassFlags = (BypassStatusFlags)_regs[7];
            }
        }

        public void QueryIPSCRStatus() //Regsiter 201 - 203
        {
            lock (_sync)
            {
                SendRequest(04, 201, 3);
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(60, 100, "QueryIPSCRStatus")));
                if (_regs == null || _regs.Length < 3) { throw new InvalidOperationException($"QueryIPSCRStatus Unexpected register length: got {_regs?.Length ?? 0}"); }
                _statusFlags4 = (SystemStatus4Flags)_regs[0];
                _statusFlags5 = (SystemStatus5Flags)_regs[1];
                _statusFlags3 = (SystemStatus3Flags)_regs[2];
            }
        }
        #endregion

        #region Control Commands
        public void ClearFaults()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA1 }, new byte[] { 0x00, 0xFF });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ClearFaults")));
            }
        }
        public void ManualBypass()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA4 }, new byte[] { 0x00, 0xFF });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualBypass")));
            }
        }
        public void ManualTransferToInverter()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA4 }, new byte[] { 0x00, 0x01 });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualTransferToInverter")));
            }
        }
        public void ECSManualBypass()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA4 }, new byte[] { 0x00, 0x02 });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ECSManualBypass")));
            }
        }

        public void BatteryTest()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA5 }, new byte[] { 0x00, 0x0F });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "BatteryTest")));
            }
        }
        public void BatteryMaintenance()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA5 }, new byte[] { 0x00, 0xF0 });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "BatteryMaintenance")));
            }
        }
        public void ManualFloat()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA5 }, new byte[] { 0xF0, 0x00 });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualFloat")));
            }
        }
        public void ManualBoost()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA5 }, new byte[] { 0x0F, 0x00 });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "ManualBoost")));
            }
        }

        public void StopTest()
        {
            lock (_sync)
            {
                SendCommand(new byte[] { 0x68, 0x28, 0xA5 }, new byte[] { 0xFF, 0xFF });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "StopTest")));
            }
        }

        public void SetBacklightTimer(int minutes)
        {
            lock (_sync)
            {
                if(minutes < 1) { minutes = 1; }
                if(minutes > 60) { minutes = 60; }
                SendCommand(new byte[] { 0x68, 0x27, 0x40 }, new byte[] { (byte)(minutes >> 8), (byte)(minutes & 0xFF) });
                _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadDataAdaptive(150, 100, "SetBacklightTimer")));
            }
        }
        #endregion
    }
}