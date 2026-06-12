using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Security.Principal;
using System.Threading;
using Microsoft.Diagnostics.Tracing;

class UsbTransaction
{
    public ulong UrbPointer { get; set; }
    public ulong DeviceHandle { get; set; }
    public ulong PipeHandle { get; set; }
    public string Type { get; set; } = "";
    public double StartTimestamp { get; set; }
    public double EndTimestamp { get; set; }
    public uint Status { get; set; }
    public ushort VendorId { get; set; }
    public ushort ProductId { get; set; }

    public double DurationUs => (EndTimestamp - StartTimestamp) * 1000.0;
    public bool IsComplete => EndTimestamp > 0;
    public string VidPid => VendorId > 0 ? $"{VendorId:X4}:{ProductId:X4}" : "";
}

class Program
{
    const string VERSION = "1.2.0";

    // Known MH devices. Gaming series is what poll checks target;
    // setup-mode PIDs are recognized only to tell the user to switch modes.
    static readonly Dictionary<string, string> KnownDevices = new()
    {
        ["39AE:400A"] = "MH4 Gamepad (Analog)",
        ["39AE:400D"] = "MH4 Gamepad (Digital)",
        ["39AE:500A"] = "MH5 Gamepad (Analog)",
        ["39AE:500D"] = "MH5 Gamepad (Digital)",
        ["054C:05C4"] = "MH Gamepad (PS4 Mode)",
        ["1A86:1235"] = "MH-XSX / XInput v1.0",
        ["39AE:4000"] = "MH4 Gamepad (Setup Mode)",
        ["39AE:5000"] = "MH5 Gamepad (Setup Mode)",
        ["1209:0001"] = "WebUSB Setup Device",
    };

    static readonly string[] SetupModeVidPids = { "39AE:4000", "39AE:5000", "1209:0001" };

    // USB mode per identity - the PID tells the mode directly.
    static readonly Dictionary<string, string> DeviceModes = new()
    {
        ["39AE:400A"] = "XInput (gaming)",
        ["39AE:400D"] = "XInput (gaming)",
        ["39AE:500A"] = "XInput (gaming)",
        ["39AE:500D"] = "XInput (gaming)",
        ["1A86:1235"] = "XInput (gaming)",
        ["054C:05C4"] = "PS4 emulation (gaming)",
        ["39AE:4000"] = "WebUSB setup",
        ["39AE:5000"] = "WebUSB setup",
        ["1209:0001"] = "WebUSB setup",
    };

    // Nominal poll rate per device; drives the "reading is below normal"
    // diagnostics gate. Diagnostics stay hidden when the reading is healthy.
    static readonly Dictionary<string, double> ExpectedRates = new()
    {
        ["39AE:400A"] = 8000,
        ["39AE:400D"] = 8000,
        ["39AE:500A"] = 8000,
        ["39AE:500D"] = 8000,
        ["054C:05C4"] = 8000,
        ["1A86:1235"] = 8000,
    };

    const string GamepadFilter = "39AE:400A,39AE:400D,39AE:500A,39AE:500D,054C:05C4,1A86:1235";

    static string DeviceDisplayName(string vidPid) =>
        KnownDevices.TryGetValue(vidPid, out var name) ? name : vidPid;

    // Raw CPU counters for measuring system load across the capture window.
    // Busy systems add genuine USB completion jitter (DPC latency); reporting
    // it preempts "why does my reading look noisy" support questions.
    static (ulong Idle, ulong Timestamp) ReadCpuRaw()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT PercentIdleTime, TimeStamp_Sys100NS FROM Win32_PerfRawData_PerfOS_Processor WHERE Name='_Total'");
            foreach (var obj in searcher.Get())
                return (Convert.ToUInt64(obj["PercentIdleTime"]), Convert.ToUInt64(obj["TimeStamp_Sys100NS"]));
        }
        catch { }
        return (0, 0);
    }

    static double CpuBusyPercent((ulong Idle, ulong Timestamp) start, (ulong Idle, ulong Timestamp) end)
    {
        if (start.Timestamp == 0 || end.Timestamp <= start.Timestamp) return -1;
        double busy = 100.0 * (1.0 - (double)(end.Idle - start.Idle) / (end.Timestamp - start.Timestamp));
        return Math.Max(0, Math.Min(100, busy));
    }

    // 1209:0001 is the shared pid.codes test PID, so VID:PID alone cannot
    // prove it is an MH board. The product string can: newer MH setup
    // firmware reports "MH-..." names. Generic strings stay unverified so we
    // never claim a random WebUSB device is an MH gamepad.
    static List<(string VidPid, string Name, bool Verified)> DetectMHDevices()
    {
        var found = new List<(string VidPid, string Name, bool Verified)>();
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT Name, PNPDeviceID FROM Win32_PnPEntity WHERE " +
                "PNPDeviceID LIKE 'USB\\\\VID_39AE%' OR PNPDeviceID LIKE 'USB\\\\VID_054C&PID_05C4%' OR " +
                "PNPDeviceID LIKE 'USB\\\\VID_1A86&PID_1235%' OR PNPDeviceID LIKE 'USB\\\\VID_1209&PID_0001%'");
            foreach (var obj in searcher.Get())
            {
                string id = obj["PNPDeviceID"]?.ToString() ?? "";
                var m = System.Text.RegularExpressions.Regex.Match(id, @"VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})");
                if (!m.Success) continue;
                string vidPid = $"{m.Groups[1].Value.ToUpper()}:{m.Groups[2].Value.ToUpper()}";
                if (found.Any(f => f.VidPid == vidPid)) continue;

                if (vidPid == "1209:0001")
                {
                    string busName = obj["Name"]?.ToString() ?? "";
                    if (busName.StartsWith("MH", StringComparison.OrdinalIgnoreCase))
                        found.Add((vidPid, $"{busName} (Setup Mode)", true));
                    else
                        found.Add((vidPid, "WebUSB Setup Device", false));
                }
                else
                {
                    found.Add((vidPid, DeviceDisplayName(vidPid), true));
                }
            }
        }
        catch { }
        return found;
    }

    const string BOX_TL = "┌";
    const string BOX_TR = "┐";
    const string BOX_BL = "└";
    const string BOX_BR = "┘";
    const string BOX_H = "─";
    const string BOX_V = "│";
    const string LINE_D = "═";

    static void Main(string[] args)
    {
        // Check admin rights for live capture
        bool isAdmin = new WindowsPrincipal(WindowsIdentity.GetCurrent())
            .IsInRole(WindowsBuiltInRole.Administrator);

        // Check if analyzing existing ETL or running live capture
        if (args.Length > 0 && File.Exists(args[0]))
        {
            // Analyze existing ETL file
            string etlPath = args[0];
            bool verbose = args.Contains("-v") || args.Contains("--verbose");
            AnalyzeEtl(etlPath, verbose);
            return;
        }

        // Live capture mode - requires admin
        if (!isAdmin)
        {
            Console.WriteLine();
            Console.WriteLine("  ╔════════════════════════════════════════════════════════╗");
            Console.WriteLine("  ║                                                        ║");
            Console.WriteLine("  ║   Right-click the exe and select 'Run as administrator' ║");
            Console.WriteLine("  ║                                                        ║");
            Console.WriteLine("  ╚════════════════════════════════════════════════════════╝");
            Console.WriteLine();
            Console.WriteLine("  Press any key to exit...");
            Console.ReadKey();
            return;
        }

        bool verbose_mode = args.Contains("-v") || args.Contains("--verbose");
        RunLiveCapture(verbose_mode);
    }

    static void RunLiveCapture(bool verbose)
    {
        Console.WriteLine();
        PrintDoubleLine(60);
        PrintCentered("D E E P P O L L", 60);
        PrintCentered($"USB Polling Analyzer  v{VERSION}", 60);
        PrintDoubleLine(60);
        Console.WriteLine();
        Console.WriteLine("  Measures USB polling rate with microsecond precision");
        Console.WriteLine("  using Windows ETW kernel tracing.");
        Console.WriteLine();
        PrintSingleLine(60);
        Console.WriteLine();
        var detected = DetectMHDevices();
        var gaming = detected.Where(d => !SetupModeVidPids.Contains(d.VidPid)).ToList();

        Console.WriteLine("  What device?");
        Console.WriteLine();
        if (gaming.Count > 0)
            Console.WriteLine($"    [1] {gaming[0].Name}  (detected)");
        else
            Console.WriteLine("    [1] MH4 / MH5 Gamepad");
        Console.WriteLine("    [2] Other USB Device");
        if (gaming.Count == 0 && detected.Count > 0)
        {
            Console.WriteLine();
            if (detected[0].Verified)
            {
                Console.WriteLine($"    Note: {detected[0].Name} found.");
                Console.WriteLine("    Calibrate it at setup.mariusheier.com first -");
                Console.WriteLine("    the board switches to Gaming mode when done.");
            }
            else
            {
                Console.WriteLine("    Note: a WebUSB device is connected. If that is your");
                Console.WriteLine("    MH board in setup mode, calibrate it at");
                Console.WriteLine("    setup.mariusheier.com - it switches to Gaming mode.");
            }
        }
        Console.WriteLine();
        Console.WriteLine("    Sending a log to support? That moved to DeepLog:");
        Console.WriteLine("    tools.mariusheier.com/deeplog");
        Console.WriteLine();
        Console.Write("  Select: ");

        string? choice = Console.ReadLine()?.Trim();
        bool isGamepad = choice == "1";

        if (!isGamepad)
        {
            // Other USB device - capture and let user select from found devices
            RunPollCheck(verbose, null);
            return;
        }

        RunPollCheck(verbose, GamepadFilter, 20);
    }

    static void RunPollCheck(bool verbose, string? deviceFilter, int durationSeconds = 7, string? deviceInstruction = null)
    {
        Console.WriteLine();
        Console.WriteLine("  ┌─────────────────────────────────────────────────────┐");
        Console.WriteLine("  │  Make sure your device is plugged in.               │");
        if (deviceInstruction != null)
        {
            Console.WriteLine("  │                                                     │");
            Console.WriteLine($"  │  {deviceInstruction,-51} │");
        }
        Console.WriteLine("  │                                                     │");
        Console.WriteLine("  │  Press ENTER to start capture.                      │");
        Console.WriteLine("  └─────────────────────────────────────────────────────┘");
        Console.ReadLine();

        string etlPath = Path.Combine(Path.GetTempPath(), "deeppoll_capture.etl");

        Console.WriteLine();
        Console.WriteLine($"  Capturing for {durationSeconds} seconds...");
        Console.WriteLine();

        var cpuTimeline = StartEtwCapture(etlPath, durationSeconds);

        Console.WriteLine("  Processing...");
        Console.WriteLine();

        AnalyzeEtl(etlPath, verbose, deviceFilter, cpuTimeline);
        try { File.Delete(etlPath); } catch { }
    }

    static List<(double TimeMs, double BusyPct)> StartEtwCapture(string etlPath, int seconds)
    {
        // Stop any existing trace
        RunCommand("logman", "stop usbpollcap -ets", silent: true);

        // Stopwatch anchored just before session start so CPU samples share the
        // trace's time axis (within command startup latency).
        var sw = System.Diagnostics.Stopwatch.StartNew();

        // Start new trace. Explicit large buffers: at 8 kHz a single device emits
        // ~16k events/s and default buffer settings silently drop events under load,
        // which reads as a lower poll rate. -ct perf pins the session clock to QPC
        // so timestamps never fall back to a coarse clock.
        RunCommand("logman", $"start usbpollcap -p Microsoft-Windows-USB-UCX -o \"{etlPath}\" -nb 64 256 -bs 512 -ct perf -ets");

        // USBHUB3 rundown events identify connected devices (VID:PID per device handle)
        RunCommand("logman", "update usbpollcap -p Microsoft-Windows-USB-USBHUB3 -ets");

        // Sample system CPU load every 500ms alongside the capture so stalls in
        // the USB data can be correlated with load spikes. Reading two counters
        // twice a second has no measurable effect on USB completion timing.
        var cpuTimeline = new List<(double TimeMs, double BusyPct)>();
        var prev = ReadCpuRaw();
        // Wall-clock based so WMI sampling time doesn't stretch the capture
        double captureStartMs = sw.Elapsed.TotalMilliseconds;
        while (sw.Elapsed.TotalMilliseconds - captureStartMs < seconds * 1000)
        {
            Thread.Sleep(500);
            var cur = ReadCpuRaw();
            double busy = CpuBusyPercent(prev, cur);
            if (busy >= 0) cpuTimeline.Add((sw.Elapsed.TotalMilliseconds, busy));
            prev = cur;
        }

        // Stop trace
        RunCommand("logman", "stop usbpollcap -ets");

        return cpuTimeline;
    }

    static int CountUsbDevices()
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                "SELECT * FROM Win32_PnPEntity WHERE PNPDeviceID LIKE 'USB%'");
            return searcher.Get().Count;
        }
        catch
        {
            return 0;
        }
    }

    static void RunCommand(string exe, string args, bool silent = false)
    {
        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName = exe,
            Arguments = args,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var proc = System.Diagnostics.Process.Start(psi);
        proc?.WaitForExit();
    }

    static void AnalyzeEtl(string etlPath, bool verbose, string? filterDevice = null,
        List<(double TimeMs, double BusyPct)>? cpuTimeline = null)
    {
        if (!File.Exists(etlPath))
        {
            Console.WriteLine("  Capture failed - no data collected.");
            return;
        }

        // Parse ETL
        var (interruptTx, controlTx, captureDurationMs, deviceBIntervals) = ParseEtl(etlPath);

        if (interruptTx.Count == 0)
        {
            Console.WriteLine("  No interrupt transfers found in trace.");
            return;
        }

        // Group by device handle to find unique devices
        var deviceGroups = interruptTx
            .GroupBy(t => t.DeviceHandle)
            .Select(g => {
                var txList = g.OrderBy(t => t.EndTimestamp).ToList();
                // Get VID:PID from first transaction that has it
                var withVidPid = txList.FirstOrDefault(t => t.VendorId > 0);
                string vidPid = withVidPid?.VidPid ?? "Unknown";

                double estimatedHz = ComputePollRate(txList).RateHz;

                return new {
                    Handle = g.Key,
                    Transactions = txList,
                    Count = g.Count(),
                    VidPid = vidPid,
                    EstimatedHz = estimatedHz
                };
            })
            .OrderByDescending(g => g.Count)
            .ToList();

        var selectedGroup = deviceGroups[0];

        // If filterDevice specified (comma-separated VID:PIDs), try to use that
        if (!string.IsNullOrEmpty(filterDevice))
        {
            var filters = filterDevice.Split(',');
            var filtered = deviceGroups.FirstOrDefault(g => filters.Contains(g.VidPid));
            // VID:PID not in ETL data - fall back to device with most samples
            if (filtered != null) selectedGroup = filtered;
        }
        // If only one device, auto-select (selectedGroup already set)
        // Multiple devices - let user select
        else if (deviceGroups.Count > 1)
        {
            Console.WriteLine();
            Console.WriteLine("  Found devices:");
            Console.WriteLine();
            for (int i = 0; i < deviceGroups.Count; i++)
            {
                var g = deviceGroups[i];
                string hzStr = g.EstimatedHz > 0 ? $"~{g.EstimatedHz:F0} Hz" : "";
                Console.WriteLine($"    [{i + 1}] {DeviceDisplayName(g.VidPid),-26} {g.Count,6:N0} samples  {hzStr}");
            }
            Console.WriteLine();
            Console.Write("  Select: ");

            string? choice = Console.ReadLine()?.Trim();
            if (!int.TryParse(choice, out int idx) || idx < 1 || idx > deviceGroups.Count)
            {
                Console.WriteLine("  Invalid selection.");
                return;
            }

            selectedGroup = deviceGroups[idx - 1];
        }

        List<UsbTransaction> selectedTransactions = selectedGroup.Transactions;
        string selectedVidPid = selectedGroup.VidPid;

        // Calculate intervals for selected device
        var sorted = selectedTransactions;
        var (pollRate, intervals, stalls, stallMs, stallThresholdUs) = ComputePollRate(sorted);
        int stallCount = stalls.Count;

        if (intervals.Count == 0)
        {
            Console.WriteLine("  Not enough data to calculate poll rate.");
            return;
        }

        var sortedIntervals = intervals.OrderBy(x => x).ToList();
        double maxInterval = sortedIntervals.Last();

        int gaps200 = intervals.Count(x => x > 200);
        int gaps500 = intervals.Count(x => x > 500);
        int gaps1ms = intervals.Count(x => x > 1000);
        int gaps5ms = intervals.Count(x => x > 5000);

        // Pretty output. Version is part of the panel because users
        // screenshot this - it tells support exactly what they ran.
        Console.WriteLine();
        PrintDoubleLine(60);
        PrintCentered("D E E P P O L L", 60);
        PrintCentered($"v{VERSION}", 60);
        PrintDoubleLine(60);

        bool isSetupMode = SetupModeVidPids.Contains(selectedVidPid);

        if (selectedVidPid != "Unknown")
            Console.WriteLine($"  Device:     {DeviceDisplayName(selectedVidPid)}  [{selectedVidPid}]");
        if (DeviceModes.TryGetValue(selectedVidPid, out var deviceMode))
            Console.WriteLine($"  Mode:       {deviceMode}");
        Console.WriteLine($"  Poll Rate:  {pollRate:F0} Hz    Samples: {intervals.Count:N0}");

        if (isSetupMode)
        {
            Console.WriteLine("  Note:       This board is in setup mode - it does not poll at");
            Console.WriteLine("              a fixed rate, so this reading says nothing about");
            Console.WriteLine("              gaming performance. Calibrate the board at");
            Console.WriteLine("              setup.mariusheier.com, then test in Gaming mode.");
        }
        // Diagnostics are only shown when the reading is actually off.
        // A healthy result stays clean - extra warnings on a good reading
        // just generate "is this ok?" support questions.
        double expectedHz = ExpectedRates.TryGetValue(selectedVidPid, out var exp) ? exp : 0;
        double spanMs = intervals.Sum() / 1000.0;
        bool offNominal = expectedHz > 0 && pollRate < expectedHz * 0.98;
        bool heavyStalls = stallCount > 0 && stallMs > spanMs * 0.005;

        if (!isSetupMode && (offNominal || heavyStalls))
        {
            if (offNominal)
                Console.WriteLine($"  Expected:   {expectedHz:F0} Hz - this reading is below normal.");

            // What the device's own USB descriptor asks for (MH boards are
            // high-speed: rate = 8000 / 2^(bInterval-1)).
            if (offNominal && deviceBIntervals.TryGetValue(selectedVidPid, out int bi) && bi > 0)
            {
                double cfgHz = 8000.0 / Math.Pow(2, bi - 1);
                Console.WriteLine($"  Configured: {cfgHz:F0} Hz (USB descriptor bInterval={bi})");
            }

            int spiked = 0;
            if (stallCount > 0)
            {
                Console.WriteLine($"  Stalls:     {stallCount} gap(s) > {stallThresholdUs / 1000.0:F1} ms excluded ({stallMs:F0} ms total)");

                // Correlate stalls with CPU spikes sampled during the capture.
                // +-1s tolerance covers the alignment slack between the trace
                // clock and the sampler's stopwatch.
                if (cpuTimeline != null && cpuTimeline.Count > 0)
                {
                    spiked = stalls.Count(s => cpuTimeline.Any(c =>
                        Math.Abs(c.TimeMs - s.AtMs) <= 1000 && c.BusyPct >= 80));
                    if (spiked > 0)
                    {
                        Console.WriteLine($"              {spiked} of {stallCount} during CPU spikes (>80% busy)");
                        Console.WriteLine("              - caused by system load, not the device.");
                    }
                }
            }

            double avgBusy = -1;
            if (cpuTimeline != null && cpuTimeline.Count > 0)
            {
                avgBusy = cpuTimeline.Average(c => c.BusyPct);
                double peakBusy = cpuTimeline.Max(c => c.BusyPct);
                if (avgBusy >= 50)
                    Console.WriteLine($"  System:     CPU {avgBusy:F0}% busy during capture (peak {peakBusy:F0}%).");
            }

            // Only blame the user's system when the capture actually implicates
            // it (busy CPU or spike-correlated stalls). If the system looks
            // fine, say nothing about it - a below-normal reading on a quiet
            // system points at the device, not their PC.
            bool systemImplicated = cpuTimeline != null && cpuTimeline.Count > 0
                ? (avgBusy >= 50 || spiked > 0)
                : stallCount > 0;
            if (systemImplicated)
            {
                Console.WriteLine("  Tip:        Background apps and power saving add jitter.");
                Console.WriteLine("              Close other apps, set the Windows power plan to");
                Console.WriteLine("              'High performance', then run the check again.");
            }
        }
        Console.WriteLine();
        Console.WriteLine("  Timing Distribution:");
        Console.WriteLine();

        // Auto-bin based on data range (zero-length intervals excluded from display)
        var distIntervals = sortedIntervals.Where(x => x > 0).ToList();
        if (distIntervals.Count == 0) distIntervals = sortedIntervals;
        double medianUs = Percentile(distIntervals, 50);
        double p5 = Percentile(distIntervals, 5);
        double p95 = Percentile(distIntervals, 95);

        // Create 8 bins covering most of the data (p5 to p95)
        double binWidth = (p95 - p5) / 6.0;
        if (binWidth < 1) binWidth = 1;

        var binList = new List<(string Label, double MinUs, double MaxUs)>();

        // First bin: everything faster than p5
        double fastestHz = 1000000.0 / p5;
        binList.Add(($">{fastestHz:F0} Hz", 0, p5));

        // Middle bins
        for (int i = 0; i < 6; i++)
        {
            double minUs = p5 + i * binWidth;
            double maxUs = p5 + (i + 1) * binWidth;
            double minHz = 1000000.0 / maxUs;
            double maxHz = 1000000.0 / minUs;
            binList.Add(($"{minHz:F0}-{maxHz:F0}", minUs, maxUs));
        }

        // Last bin: everything slower than p95
        double slowestHz = 1000000.0 / p95;
        binList.Add(($"<{slowestHz:F0} Hz", p95, double.MaxValue));

        var bins = binList.ToArray();

        int maxCount = 0;
        var binCounts = new int[bins.Length];
        for (int i = 0; i < bins.Length; i++)
        {
            binCounts[i] = intervals.Count(x => x >= bins[i].MinUs && x < bins[i].MaxUs);
            if (binCounts[i] > maxCount) maxCount = binCounts[i];
        }

        for (int i = 0; i < bins.Length; i++)
        {
            double pct = 100.0 * binCounts[i] / intervals.Count;
            int barLen = maxCount > 0 ? (int)(35.0 * binCounts[i] / maxCount) : 0;

            char barChar = (i == 0 || i == bins.Length - 1) ? '░' : '▓';
            string bar = new string(barChar, barLen);

            Console.WriteLine($"  {bins[i].Label,12} {bar.PadRight(35)} {pct,5:F1}%");
        }

        PrintDoubleLine(60);
        Console.WriteLine();

        // Verbose output for diagnostics (your eyes only)
        if (verbose)
        {
            PrintVerboseAnalysis(interruptTx, controlTx, intervals, sortedIntervals,
                gaps200, gaps500, gaps1ms, gaps5ms, maxInterval);

            if (deviceBIntervals.Count > 0)
            {
                Console.WriteLine("\nDEVICE DESCRIPTORS (fastest interrupt IN)");
                PrintSingleLine(60);
                foreach (var kv in deviceBIntervals)
                {
                    string rate = kv.Value > 0
                        ? $"bInterval={kv.Value} ({8000.0 / Math.Pow(2, kv.Value - 1):F0} Hz at high speed)"
                        : "no interrupt IN endpoint (bulk/setup interface)";
                    Console.WriteLine($"  {kv.Key}  {rate}");
                }
            }
        }
    }

    static (List<UsbTransaction> interrupt, List<UsbTransaction> control, double durationMs,
        Dictionary<string, int> deviceBIntervals) ParseEtl(string etlPath)
    {
        var pendingTransactions = new Dictionary<ulong, UsbTransaction>();
        var completedTransactions = new List<UsbTransaction>();
        var deviceVidPids = new Dictionary<ulong, (ushort Vid, ushort Pid)>();
        // VID:PID -> fastest interrupt IN bInterval from the config descriptor
        // (0 = config has no interrupt IN endpoint; absent = descriptor unseen)
        var deviceBIntervals = new Dictionary<string, int>();
        double firstTimestamp = -1;
        double lastTimestamp = 0;

        using (var source = new ETWTraceEventSource(etlPath))
        {
            source.Dynamic.All += (TraceEvent data) =>
            {
                string provider = data.ProviderName ?? "";
                string eventName = data.EventName ?? "";

                if (provider.Contains("USBHUB3"))
                {
                    // Hub rundown/device events carry the same fid_UsbDevice handle
                    // that UCX transfer events are grouped by, plus the device path
                    // with VID/PID. This is the only reliable handle->device mapping.
                    ulong hubUsbDevice = 0;
                    string devicePath = "";
                    try { hubUsbDevice = Convert.ToUInt64(data.PayloadByName("fid_UsbDevice")); } catch { }
                    try { devicePath = data.PayloadByName("fid_DeviceInterfacePath")?.ToString() ?? ""; } catch { }
                    if (hubUsbDevice != 0 && devicePath.Length > 0)
                    {
                        var m = System.Text.RegularExpressions.Regex.Match(devicePath, @"VID_([0-9A-Fa-f]{4})&PID_([0-9A-Fa-f]{4})");
                        if (m.Success)
                        {
                            ushort vid = Convert.ToUInt16(m.Groups[1].Value, 16);
                            ushort pid = Convert.ToUInt16(m.Groups[2].Value, 16);
                            deviceVidPids[hubUsbDevice] = (vid, pid);

                            byte[]? cfg = null;
                            try { cfg = data.PayloadByName("fid_ConfigurationDescriptor") as byte[]; } catch { }
                            if (cfg != null && cfg.Length > 0)
                                deviceBIntervals[$"{vid:X4}:{pid:X4}"] = FastestInterruptInBInterval(cfg);
                        }
                    }
                    return;
                }

                if (!provider.Contains("USB-UCX")) return;

                bool isDispatch = eventName.Contains("/Start");
                bool isComplete = eventName.Contains("/Stop");
                if (!isDispatch && !isComplete) return;

                ulong urbPtr = 0;
                try { urbPtr = Convert.ToUInt64(data.PayloadByName("fid_URB_Ptr")); } catch { }
                if (urbPtr == 0) return;

                string txType = "Unknown";
                if (eventName.Contains("BULK_OR_INTERRUPT")) txType = "Interrupt";
                else if (eventName.Contains("CONTROL") || eventName.Contains("CLASS_INTERFACE")) txType = "Control";

                if (firstTimestamp < 0) firstTimestamp = data.TimeStampRelativeMSec;
                lastTimestamp = data.TimeStampRelativeMSec;

                if (isDispatch)
                {
                    ulong deviceHandle = 0;
                    ulong pipeHandle = 0;
                    ushort vid = 0, pid = 0;
                    try { deviceHandle = Convert.ToUInt64(data.PayloadByName("fid_UsbDevice")); } catch { }
                    try { pipeHandle = Convert.ToUInt64(data.PayloadByName("fid_PipeHandle")); } catch { }
                    try { vid = Convert.ToUInt16(data.PayloadByName("fid_idVendor")); } catch { }
                    try { pid = Convert.ToUInt16(data.PayloadByName("fid_idProduct")); } catch { }

                    pendingTransactions[urbPtr] = new UsbTransaction
                    {
                        UrbPointer = urbPtr,
                        DeviceHandle = deviceHandle,
                        PipeHandle = pipeHandle,
                        Type = txType,
                        StartTimestamp = data.TimeStampRelativeMSec,
                        VendorId = vid,
                        ProductId = pid
                    };
                }
                else if (isComplete)
                {
                    if (pendingTransactions.TryGetValue(urbPtr, out var tx))
                    {
                        tx.EndTimestamp = data.TimeStampRelativeMSec;
                        try { tx.Status = Convert.ToUInt32(data.PayloadByName("fid_IRP_NtStatus")); } catch { }

                        if (tx.DurationUs < 100000)
                            completedTransactions.Add(tx);

                        pendingTransactions.Remove(urbPtr);
                    }
                }
            };

            source.Process();
        }

        // Fill in VID:PID from the hub device map (transfer events don't carry it)
        foreach (var tx in completedTransactions)
        {
            if (tx.VendorId == 0 && deviceVidPids.TryGetValue(tx.DeviceHandle, out var vp))
            {
                tx.VendorId = vp.Vid;
                tx.ProductId = vp.Pid;
            }
        }

        var interrupt = completedTransactions.Where(t => t.Type == "Interrupt").ToList();
        var control = completedTransactions.Where(t => t.Type == "Control").ToList();

        return (interrupt, control, lastTimestamp - firstTimestamp, deviceBIntervals);
    }

    static void PrintVerboseAnalysis(List<UsbTransaction> interruptTx, List<UsbTransaction> controlTx,
        List<double> intervals, List<double> sortedIntervals,
        int gaps200, int gaps500, int gaps1ms, int gaps5ms, double maxInterval)
    {
        Console.WriteLine("DIAGNOSTIC DATA (not shown to customer)");
        PrintDoubleLine(60);

        // Gap counts
        Console.WriteLine("\nGAP COUNTS");
        PrintSingleLine(60);
        int maxGap = Math.Max(1, new[] { gaps200, gaps500, gaps1ms, gaps5ms }.Max());
        PrintGapBar(">200μs", gaps200, maxGap, 40);
        PrintGapBar(">500μs", gaps500, maxGap, 40);
        PrintGapBar(">1ms", gaps1ms, maxGap, 40);
        PrintGapBar(">5ms", gaps5ms, maxGap, 40);
        Console.WriteLine($"\n  Max interval: {maxInterval:F1} μs");

        // URB durations
        var durations = interruptTx.Select(t => t.DurationUs).OrderBy(x => x).ToList();
        Console.WriteLine("\nURB DURATION (Dispatch → Complete)");
        PrintSingleLine(60);
        Console.WriteLine($"  Samples: {durations.Count:N0}");
        Console.WriteLine($"  P50:     {Percentile(durations, 50):F1} μs");
        Console.WriteLine($"  P99:     {Percentile(durations, 99):F1} μs");
        Console.WriteLine($"  P99.9:   {Percentile(durations, 99.9):F1} μs");
        Console.WriteLine($"  Max:     {durations.Last():F1} μs");

        // Largest gaps
        var sorted = interruptTx.OrderBy(t => t.EndTimestamp).ToList();
        var gapsWithTime = new List<(double Interval, double Timestamp)>();
        for (int i = 1; i < sorted.Count; i++)
        {
            double interval = (sorted[i].EndTimestamp - sorted[i - 1].EndTimestamp) * 1000;
            if (interval > 200)
                gapsWithTime.Add((interval, sorted[i].EndTimestamp));
        }

        if (gapsWithTime.Count > 0)
        {
            Console.WriteLine("\nLARGEST GAPS");
            PrintSingleLine(60);
            foreach (var gap in gapsWithTime.OrderByDescending(g => g.Interval).Take(10))
            {
                Console.WriteLine($"  {gap.Interval,10:F1} μs  at  {gap.Timestamp:F3} ms");
            }
        }

        // EP0 correlation
        if (controlTx.Count > 0)
        {
            Console.WriteLine("\nEP0 CONTROL TRANSFERS");
            PrintSingleLine(60);
            Console.WriteLine($"  Count: {controlTx.Count}");

            foreach (var ep0 in controlTx.Take(5))
            {
                Console.WriteLine($"\n  EP0 at {ep0.StartTimestamp:F3}ms (duration: {ep0.DurationUs:F0}μs)");

                var nearby = interruptTx
                    .Where(t => Math.Abs(t.EndTimestamp - ep0.EndTimestamp) < 2)
                    .OrderBy(t => t.EndTimestamp)
                    .ToList();

                foreach (var tx in nearby.Take(10))
                {
                    double delta = (tx.EndTimestamp - ep0.EndTimestamp) * 1000;
                    string marker = Math.Abs(delta) < 100 ? " ◄" : "";
                    Console.WriteLine($"    {delta:+000;-000}μs: {tx.DurationUs:F0}μs{marker}");
                }
            }
        }

        Console.WriteLine();
    }

    // Walk a USB configuration descriptor and return the fastest interrupt IN
    // endpoint's bInterval. Returns 0 if the config has no interrupt IN
    // endpoint at all (e.g. WebUSB setup mode, which is bulk-only - readings
    // from such an interface say nothing about gaming poll rate).
    static int FastestInterruptInBInterval(byte[] cfg)
    {
        int best = 0;
        int i = 0;
        while (i + 1 < cfg.Length)
        {
            int len = cfg[i];
            if (len < 2) break;
            if (cfg[i + 1] == 0x05 && len >= 7 && i + 6 < cfg.Length)  // endpoint descriptor
            {
                bool isIn = (cfg[i + 2] & 0x80) != 0;
                bool isInterrupt = (cfg[i + 3] & 0x03) == 0x03;
                int interval = cfg[i + 6];
                if (isIn && isInterrupt && interval > 0 && (best == 0 || interval < best))
                    best = interval;
            }
            i += len;
        }
        return best;
    }

    // Steady-state poll rate from completion intervals.
    // - Zero-length intervals are real completions (batched DPC delivery on some
    //   xHCI/timer configurations) and must be counted, not filtered.
    // - Host-side stalls (DPC storms, driver warmup) pause the whole bus for
    //   milliseconds; excluding them keeps the headline number at what the
    //   device actually sustains. They are reported separately.
    static (double RateHz, List<double> Intervals, List<(double IntervalUs, double AtMs)> Stalls, double StallMs, double ThresholdUs)
        ComputePollRate(List<UsbTransaction> txSorted)
    {
        var all = new List<double>();
        var timed = new List<(double IntervalUs, double AtMs)>();
        for (int i = 1; i < txSorted.Count; i++)
        {
            double interval = (txSorted[i].EndTimestamp - txSorted[i - 1].EndTimestamp) * 1000;
            if (interval >= 0 && interval < 50000)
            {
                all.Add(interval);
                timed.Add((interval, txSorted[i].EndTimestamp));
            }
        }
        if (all.Count == 0) return (0, all, timed, 0, 0);

        double median = Percentile(all.OrderBy(x => x).ToList(), 50);
        double thresholdUs = Math.Max(4 * median, 1000);

        var steady = all.Where(x => x < thresholdUs).ToList();
        var stalls = timed.Where(t => t.IntervalUs >= thresholdUs).ToList();
        double stallMs = stalls.Sum(t => t.IntervalUs) / 1000.0;

        double avgUs = steady.Count > 0 ? steady.Average() : 0;
        double rate = avgUs > 0 ? 1000000.0 / avgUs : 0;
        return (rate, all, stalls, stallMs, thresholdUs);
    }

    static double Percentile(List<double> sorted, double p)
    {
        if (sorted.Count == 0) return 0;
        int idx = Math.Min((int)(sorted.Count * p / 100), sorted.Count - 1);
        return sorted[idx];
    }

    static void PrintDoubleLine(int width)
    {
        Console.WriteLine(new string('═', width));
    }

    static void PrintSingleLine(int width, string? title = null)
    {
        if (title == null)
        {
            Console.WriteLine(new string('─', width));
        }
        else
        {
            int padding = (width - title.Length - 2) / 2;
            Console.WriteLine(new string('─', padding) + " " + title + " " + new string('─', width - padding - title.Length - 2));
        }
    }

    static void PrintCentered(string text, int width)
    {
        int padding = (width - text.Length) / 2;
        Console.WriteLine(new string(' ', padding) + text);
    }

    static void PrintGapBar(string label, int count, int maxCount, int barWidth)
    {
        int barLen = maxCount > 0 ? (int)((double)count / maxCount * barWidth) : 0;
        barLen = Math.Max(count > 0 ? 1 : 0, barLen);

        string bar = new string('█', barLen);
        Console.WriteLine($"  {label,7}  {bar.PadRight(barWidth)}  {count,6:N0}");
    }
}
