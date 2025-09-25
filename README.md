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
            Console.WriteLine($"Model: {ups.Model}, Version: {ups.Version}, Batteries: {ups.BatteryCount}");

            // Query full status
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

## Manufacturer & Firmware

| Property       | Type   | Description                                     |
| -------------- | ------ | ----------------------------------------------- |
| `Model`        | string | UPS model name retrieved from manufacturer info |
| `Version`      | string | Firmware version                                |
| `BatteryCount` | int    | Number of batteries connected                   |

---

## Power Metrics

| Property             | Type  | Description                                |
| -------------------- | ----- | ------------------------------------------ |
| `BypassVoltage`      | float | Bypass input voltage (V)                   |
| `MainVoltage`        | float | Main input voltage (V)                     |
| `OutputVoltage`      | float | Output voltage (V)                         |
| `VA`                 | float | Apparent power (VA)                        |
| `Watts`              | float | Active power (W)                           |
| `BatteryVoltage`     | float | Battery voltage (V)                        |
| `BatteryCurrent`     | float | Battery current (A)                        |
| `BatteryCapacity`    | float | Battery charge percentage (%)              |
| `BatteryTimeRemain`  | float | Estimated battery time remaining (minutes) |
| `BatteryTemp`        | float | Battery temperature (°C)                   |
| `Runtime`            | float | UPS runtime (minutes)                      |
| `BusVoltage`         | float | Internal bus voltage (V)                   |
| `Syscode`            | float | System code (internal)                     |
| `LoadPercent`        | float | Load percentage (%)                        |
| `BypassFreq`         | float | Bypass input frequency (Hz)                |
| `MainFreq`           | float | Main input frequency (Hz)                  |
| `OutputFreq`         | float | Output frequency (Hz)                      |
| `RatedOutputVoltage` | float | Rated output voltage (V)                   |
| `RatedOutputFreq`    | float | Rated output frequency (Hz)                |
| `RatedInputVoltage`  | float | Rated input voltage (V)                    |
| `RatedInputFreq`     | float | Rated input frequency (Hz)                 |
| `MaxWatt`            | float | Maximum output wattage (W)                 |
| `MainCurrent`        | float | Main input current (A)                     |
| `OutCurrent`         | float | Output current (A)                         |
| `BypassPF`           | float | Bypass input power factor                  |
| `MainPF`             | float | Main input power factor                    |
| `OutputPF`           | float | Output power factor                        |

---

## Charger & Inverter Info

| Property          | Type  | Description                     |
| ----------------- | ----- | ------------------------------- |
| `ChargerCurrent`  | float | Charger current (A)             |
| `InverterVoltage` | float | Inverter voltage (V)            |
| `RecIGBTTemp`     | float | Rectifier IGBT temperature (°C) |
| `InvIGBTTemp`     | float | Inverter IGBT temperature (°C)  |

---

## Switch & Output Status

| Property              | Type         | Description                     |
| --------------------- | ------------ | ------------------------------- |
| `CurrentSwitchState`  | SwitchState  | Current switch state of the UPS |
| `CurrentOutputStatus` | OutputStatus | Output status information       |
| `CurrentHardware`     | HardwareType | Detected UPS hardware type      |
| `BatteryConnected`    | bool         | Battery connection status       |

---

## System Status Flags

| Property                | Type               | Description                                   |
| ----------------------- | ------------------ | --------------------------------------------- |
| `StatusFlags`           | SystemStatusFlags  | Primary system status flags                   |
| `StatusFlags2`          | SystemStatus2Flags | Secondary system status flags                 |
| `IsAmbientOverTemp`     | bool               | True if ambient temperature exceeds threshold |
| `IsRecCanFail`          | bool               | True if rectifier CAN bus has failed          |
| `IsInvIoCanFail`        | bool               | True if inverter IO CAN bus has failed        |
| `IsInvDataCanFail`      | bool               | True if inverter data CAN bus has failed      |
| `IsBypassPowerFuseFail` | bool               | True if bypass fuse has failed                |
| `IsRatedKvaOverRange`   | bool               | True if rated kVA exceeds range               |

---

## Internal Buffers

| Property         | Type   | Description                               |
| ---------------- | ------ | ----------------------------------------- |
| `BufferUnderRun` | int    | Internal buffer under-run counter         |
| `UnderRunMethod` | string | Method or status causing buffer under-run |

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
