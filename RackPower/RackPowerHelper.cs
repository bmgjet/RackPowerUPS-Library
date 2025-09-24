using System;
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
                throw new ArgumentException("Frame too short");

            ushort rawCrc = (ushort)(raw[raw.Length - 2] | (raw[raw.Length - 1] << 8));
            ushort calc = Crc16(raw, raw.Length - 2);

            var frame = new ModbusFrame
            {
                SlaveAddress = raw[0],
                FunctionCode = raw[1],
                Crc = rawCrc,
                CrcValid = (rawCrc == calc)
            };

            int dataLen = raw.Length - 4;
            if (dataLen > 0)
            {
                frame.Data = new byte[dataLen];
                Array.Copy(raw, 2, frame.Data, 0, dataLen);
            }
            return frame;
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

        public static ushort[] ExtractRegisters(ModbusFrame frame)
        {
            if (frame == null) throw new ArgumentNullException(nameof(frame));
            if (frame.Data == null || frame.Data.Length == 0) return new ushort[0];

            int byteCount = frame.Data[0];
            if (byteCount <= 0 || byteCount != frame.Data.Length - 1)
                byteCount = frame.Data.Length;

            int offset = (frame.Data.Length > 0 && frame.Data[0] == byteCount) ? 1 : 0;
            int regCount = (frame.Data.Length - offset) / 2;

            var regs = new ushort[regCount];
            for (int i = 0; i < regCount; i++)
            {
                int idx = offset + i * 2;
                regs[i] = (ushort)((frame.Data[idx] << 8) | frame.Data[idx + 1]);
            }
            return regs;
        }

        public static string ByteToString(byte[] regs, int regIndex, int lengthRegs)
        {
            if (regs == null) { throw new ArgumentNullException(nameof(regs)); }
            if (regIndex < 0 || regIndex + lengthRegs > regs.Length) {throw new IndexOutOfRangeException(); }
            return Encoding.ASCII.GetString(regs, regIndex, lengthRegs);
        }
    }

    public class RackPowerUpsClient : IDisposable
    {
        public SerialPort _serialPort;
        public ushort[] _regs;

        // UPS Values
        public string Model { get; private set; }
        public string Version { get; private set; }
        public int BatteryCount { get; private set; }

        public float BypassVoltage { get; private set; }
        public float MainVoltage { get; private set; }
        public float OutputVoltage { get; private set; }
        public float VA { get; private set; }
        public float Watts { get; private set; }
        public float BatteryVoltage { get; private set; }
        public float BatteryCurrent { get; private set; }
        public float BatteryCapacity { get; private set; }
        public float BatteryTimeRemain { get; private set; }
        public float BatteryTemp { get; private set; }
        public float Runtime { get; private set; }
        public float BusVoltage { get; private set; }
        public float Syscode { get; private set; }
        public float LoadPercent { get; private set; }
        public float BypassFreq { get; private set; }
        public float MainFreq { get; private set; }
        public float OutputFreq { get; private set; }
        public float RatedOutputVoltage { get; private set; }
        public float RatedFreq { get; private set; }
        public float RatedInputVoltage { get; private set; }
        public float RatedInputFreq { get; private set; }
        public float ChargerCurrent { get; private set; }
        public float InverterVoltage { get; private set; }
        public float RecIGBTTemp { get; private set; }
        public float InvIGBTTemp { get; private set; }
        public float MaxWatt { get; private set; }
        public float MainCurrent { get; private set; }
        public float OutCurrent { get; private set; }

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
        }

        // === Queries ===
        public void QueryManufacturerInfo()
        {
            SendData("010327110065df50");
            Thread.Sleep(250);
            var frame = ModbusHelper.ParseFrame(ReadData());
            Model = ModbusHelper.ByteToString(frame.Data, 143, 8);
            _regs = ModbusHelper.ExtractRegisters(frame);
            BatteryCount = _regs[10];

            // Version + Temps
            SendData("01044ea3000b56c7");
            Thread.Sleep(50);
            _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadData()));
            Version = _regs[7].ToString() + "." + _regs[8] + "." + _regs[9];
            RecIGBTTemp = _regs[0] * 0.1f;
            InvIGBTTemp = _regs[3] * 0.1f;
        }

        public void QueryFullStatus()
        {
            // Power Info
            SendData("01044e210067f6c2");
            Thread.Sleep(250);
            _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadData()));
            BypassVoltage = _regs[0] * 0.1f;
            MainVoltage = _regs[12] * 0.1f;
            OutputVoltage = _regs[24] * 0.1f;
            VA = _regs[36];
            Watts = _regs[39];
            BatteryVoltage = _regs[49] * 0.1f;
            BatteryCurrent = _regs[51] * 0.1f;
            BatteryCapacity = _regs[55] * 0.1f;
            BatteryTemp = _regs[53] * 0.1f;
            BatteryTimeRemain = _regs[54] * 0.1f;
            Runtime = _regs[69] * 0.1f;
            BusVoltage = _regs[102] * 0.1f;
            MaxWatt = _regs[100];
            Syscode = _regs[93];
            LoadPercent = _regs[45] * 0.1f;
            BypassFreq = _regs[6] * 0.01f;
            MainFreq = _regs[18] * 0.01f;
            OutputFreq = _regs[30] * 0.01f;
            RatedOutputVoltage = _regs[97];
            RatedFreq = _regs[98];
            RatedInputVoltage = _regs[95];
            RatedInputFreq = _regs[96];
            MainCurrent = _regs[15] * 0.1f;
            OutCurrent = _regs[27] * 0.1f;

            // Charger/Inverter Info
            SendData("01044e8b0003d709");
            Thread.Sleep(25);
            _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadData()));
            ChargerCurrent = _regs[0] * 0.1f;
            InverterVoltage = _regs[2] * 0.1f;

            // Temps
            SendData("01044ea3000b56c7");
            Thread.Sleep(50);
            _regs = ModbusHelper.ExtractRegisters(ModbusHelper.ParseFrame(ReadData()));
            RecIGBTTemp = _regs[0] * 0.1f;
            InvIGBTTemp = _regs[3] * 0.1f;
        }

        // === Control Commands ===
        public void ClearFaults() => SendData("016828A100FF7801");
        public void ManualBypass() => SendData("016828A400FF6800");
        public void ManualTransferToInverter() => SendData("016828A40001E980");
        public void ECSManualBypass() => SendData("016828A40002A981");

        public void BatteryTest() => SendData("016828A5000F3984");
        public void BatteryMaintenance() => SendData("016828A500F079C4");
        public void ManualFloat() => SendData("016828A5F0003D80");
        public void ManualBoost() => SendData("016828A50F007C70");
        public void StopTest() => SendData("016828A5FFFF7830");

        public void SetBacklightTimer(int minutes)
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
        }

        // === Internal Helpers ===
        private void SendData(string hexString)
        {
            byte[] buffer = Enumerable.Range(0, hexString.Length / 2).Select(x => Convert.ToByte(hexString.Substring(x * 2, 2), 16)).ToArray();
            _serialPort.Write(buffer, 0, buffer.Length);
        }

        private byte[] ReadData()
        {
            int bytesToRead = _serialPort.BytesToRead;
            if (bytesToRead > 0)
            {
                byte[] response = new byte[bytesToRead];
                _serialPort.Read(response, 0, response.Length);
                return response;
            }
            return null;
        }
    }
}
