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
