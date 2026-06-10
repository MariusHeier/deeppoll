# DeepPoll

USB polling rate analyzer with kernel-level ETW tracing.

## What is this?

DeepPoll measures your USB device's actual polling rate at the kernel level using Event Tracing for Windows (ETW). Unlike browser-based tools that measure JavaScript event timing, DeepPoll captures the raw USB interrupt timing as seen by the Windows kernel.

## Quick Start

Run in PowerShell (as Administrator):

```powershell
irm tools.mariusheier.com/deeppoll.ps1 | iex
```

This downloads and runs the DeepPoll executable with checksum verification.

## Why Kernel-Level?

Browser-based polling analyzers have limitations:
- JavaScript timer resolution (~4ms minimum)
- Browser event coalescing
- OS scheduling delays
- Only sees events that reach the browser

DeepPoll uses ETW to capture USB interrupts directly from the kernel, giving you:
- Microsecond precision timing
- True interrupt distribution analysis
- Detection of missed polls and jitter
- Accurate measurement up to 8000Hz

## Supported MH Devices

DeepPoll auto-detects MH gamepads (Gaming mode):

| Device | Analog    | Digital   |
|--------|-----------|-----------|
| MH4    | 39AE:400A | 39AE:400D |
| MH5    | 39AE:500A | 39AE:500D |
| MH-XSX / XInput v1.0 firmware | 1A86:1235 | — |

Legacy MH4 units (054C:05C4) are also recognized.

Setup mode (39AE:4000 / 39AE:5000 / 1209:0001 WebUSB) is recognized but
not used for poll checks — calibrate the board at setup.mariusheier.com
first; it switches to Gaming mode when done. Any other USB device can
be measured via the "Other USB Device" option.

## Requirements

- Windows 10/11
- Administrator privileges (required for ETW tracing)
- .NET 8.0 Runtime (or use the self-contained exe)

## Verification

Every release includes SHA256 checksums. The bootstrap script automatically verifies the download.

Manual verification:
```powershell
# Download checksum
$checksum = (Invoke-WebRequest "https://github.com/MariusHeier/deeppoll/releases/latest/download/checksums.txt").Content

# Verify
Get-FileHash DeepPoll.exe -Algorithm SHA256
```

## Building from Source

```bash
cd DeepPoll
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true
```

## License

MIT License - See [LICENSE](LICENSE)

## Links

- [Web Page](https://tools.mariusheier.com/deeppoll)
- [Releases](https://github.com/MariusHeier/deeppoll/releases)
