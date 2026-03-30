# DualSenseMapper

DualSenseMapper is a lightweight Windows app that maps a Sony DualSense controller to a virtual Xbox 360 controller.

## Preview
<img width="593" height="421" alt="ui" src="https://github.com/user-attachments/assets/996b5c6b-38d6-4aba-aff6-c6cdd782f07a" />

## Features

- DualSense support over USB and Bluetooth
- Low-latency HID input setup
- Xbox 360 virtual controller output through ViGEmBus
- Mapping for sticks, triggers, face buttons, D-pad, Options, Share, PS, and touchpad click
- DualSense lightbar hue slider in the app UI


## Requirements

- Windows
- .NET 8 SDK
- ViGEmBus installed

## Run

1. Build and run from the project:

```powershell
dotnet run --project .\\DualSenseMapper.csproj
```

Or build it:

```powershell
dotnet build .\\DualSenseMapper.csproj
```

2. Download from release:  [Latest Release](../../releases/latest)

## Files

- `AppMain.cs` - app UI and controller mapping logic
- `DualSenseMapper.csproj` - project file
- `DualSenseMapper.sln` - Visual Studio solution

## Notes

- The repo ignores generated output folders like `bin`, `obj`, `build-check`, and `DualSenseMapper-App`.
- The virtual controller appears as an Xbox 360 controller, which is expected for this project.
