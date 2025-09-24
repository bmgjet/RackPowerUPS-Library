# RackPower UPS C# Library

A simple C# .NET Framework 4.7+ library for interfacing with **RackPower UPS** devices over serial (RS-232 / USB).  
Implements Modbus RTU communication for reading status, querying info, and sending control commands.

---

## ✨ Features
- ✅ Parse Modbus RTU frames with CRC validation  
- ✅ Extract UPS registers and convert to engineering units  
- ✅ Query full UPS status (voltages, currents, frequencies, temps, runtime, load, etc.)  
- ✅ Query manufacturer & firmware info  
- ✅ Control commands:
  - Clear faults  
  - Manual bypass / inverter transfer  
  - Battery test, maintenance, float, boost  
  - Backlight timer settings  

---

## 🚀 Usage Example

```csharp
using System;
using RackPowerUPS;

class Program
{
    static void Main()
    {
        using (var ups = new RackPowerUpsClient("COM3", 9600))
        {
            ups.Open();

            // Query manufacturer info
            ups.QueryManufacturerInfo();
            Console.WriteLine($"Model: {ups.Model}, Version: {ups.Version}, Batteries: {ups.BatteryCount}");

            // Query full status
            ups.QueryFullStatus();
            Console.WriteLine($"Input Voltage: {ups.MainVoltage} V");
            Console.WriteLine($"Output Voltage: {ups.OutputVoltage} V");
            Console.WriteLine($"Load: {ups.Watts} W ({ups.LoadPercent}%)");
            Console.WriteLine($"Battery: {ups.BatteryVoltage} V {ups.BatteryCapacity}%");

            ups.Close();
        }
    }
}
```
## ⚡ Available Commands
Queries

- QueryManufacturerInfo()
  Reads UPS model, version, battery count, and temps

- QueryFullStatus()
 Reads voltages, frequencies, load, runtime, temps, currents, etc.

- QueryAutoDetect()

- QueryPowerInfo()

- QueryChargerInverterInfo()

- QueryTemps()

- QueryVersionTemps()

 Control

- ClearFaults()

- ManualBypass()

- ManualTransferToInverter()

- ECSManualBypass()

- BatteryTest()

- BatteryMaintenance()

- ManualFloat()

- ManualBoost()

- StopTest()

- SetBacklightTimer(minutes) → Valid values: 1, 3, 5, 10, 20, 30

## 🛠️ Technical Details

Protocol: Modbus RTU

Default Serial: 9600 8N1 (customizable in constructor)

CRC16 implementation (0xA001 polynomial)

Uses System.IO.Ports.SerialPort

## 📄 License

MIT License – free to use, modify, and distribute.
