using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Security.Principal;
using System.Text.Json;
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
        PrintCentered("USB Polling Analyzer", 60);
        PrintDoubleLine(60);
        Console.WriteLine();
        Console.WriteLine("  Measures USB polling rate with microsecond precision");
        Console.WriteLine("  using Windows ETW kernel tracing.");
        Console.WriteLine();
        PrintSingleLine(60);
        Console.WriteLine();
        Console.WriteLine("  What device?");
        Console.WriteLine();
        Console.WriteLine("    [1] MH4 Gamepad");
        Console.WriteLine("    [2] Other USB Device");
        Console.WriteLine();
        Console.Write("  Select: ");

        string? choice = Console.ReadLine()?.Trim();
        bool isMH4 = choice == "1";

        if (!isMH4)
        {
            // Other USB device - capture and let user select from found devices
            RunPollCheck(verbose, null);
            return;
        }

        // MH4 - more options
        Console.WriteLine();
        PrintSingleLine(60);
        Console.WriteLine();
        Console.WriteLine("  What do you want to do?");
        Console.WriteLine();
        Console.WriteLine("    [1] Check Poll Rate");
        Console.WriteLine("    [2] Log USB for Marius (support)");
        Console.WriteLine();
        Console.Write("  Select: ");

        choice = Console.ReadLine()?.Trim();

        if (choice == "1")
        {
            RunPollCheck(verbose, "054C:05C4", 20);
            return;
        }

        // Log for Marius
        Console.WriteLine();
        PrintSingleLine(60);
        Console.WriteLine();
        Console.WriteLine("  What issue are you logging?");
        Console.WriteLine();
        Console.WriteLine("    [1] Startup / Connection issues (20 sec)");
        Console.WriteLine("    [2] Disconnect during gameplay (up to 30 min)");
        Console.WriteLine();
        Console.Write("  Select: ");

        choice = Console.ReadLine()?.Trim();

        if (choice == "1")
        {
            RunStartupLog();
        }
        else
        {
            RunDisconnectLog();
        }
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

        StartEtwCapture(etlPath, durationSeconds);

        Console.WriteLine("  Processing...");
        Console.WriteLine();

        AnalyzeEtl(etlPath, verbose, deviceFilter);
        try { File.Delete(etlPath); } catch { }
    }

    static void RunStartupLog()
    {
        Console.WriteLine();
        Console.WriteLine("  ┌─────────────────────────────────────────────────────┐");
        Console.WriteLine("  │  UNPLUG your MH4 controller now.                    │");
        Console.WriteLine("  │                                                     │");
        Console.WriteLine("  │  Press ENTER when ready, then plug it in.           │");
        Console.WriteLine("  └─────────────────────────────────────────────────────┘");
        Console.ReadLine();

        string etlPath = Path.Combine(Path.GetTempPath(), "deeppoll_startup.etl");

        Console.WriteLine();
        Console.WriteLine("  Plug in now! Capturing for 20 seconds...");
        Console.WriteLine();

        StartEtwCapture(etlPath, 20);

        Console.WriteLine("  Processing...");
        Console.WriteLine();

        AnalyzeEtl(etlPath, false, "054C:05C4");

        string? logId = UploadLog(etlPath, "startup");
        if (logId != null)
        {
            Console.WriteLine();
            Console.WriteLine($"  ┌─────────────────────────────────────────────────────┐");
            Console.WriteLine($"  │  Upload complete!                                    │");
            Console.WriteLine($"  │                                                     │");
            Console.WriteLine($"  │  Log ID: {logId,-43} │");
            Console.WriteLine($"  │                                                     │");
            Console.WriteLine($"  │  Send this ID to Marius on Discord:                 │");
            Console.WriteLine($"  │  https://discord.gg/4Q9SRUt85j                      │");
            Console.WriteLine($"  └─────────────────────────────────────────────────────┘");
        }

        try { File.Delete(etlPath); } catch { }
    }

    static void RunDisconnectLog()
    {
        // Check disk space (need ~500MB free for 30 min capture)
        var drive = new DriveInfo(Path.GetTempPath());
        long freeGB = drive.AvailableFreeSpace / 1024 / 1024 / 1024;
        if (freeGB < 1)
        {
            Console.WriteLine();
            Console.WriteLine("  Not enough disk space! Need at least 1GB free.");
            Console.WriteLine($"  Available: {freeGB}GB");
            return;
        }

        Console.WriteLine();
        Console.WriteLine("  ┌─────────────────────────────────────────────────────┐");
        Console.WriteLine("  │  Make sure your MH4 is connected.                   │");
        Console.WriteLine("  │                                                     │");
        Console.WriteLine("  │  Press ENTER to start logging.                      │");
        Console.WriteLine("  │  Press ENTER again when disconnect happens.         │");
        Console.WriteLine("  │  (Max 30 minutes)                                   │");
        Console.WriteLine("  └─────────────────────────────────────────────────────┘");
        Console.ReadLine();

        string etlPath = Path.Combine(Path.GetTempPath(), "deeppoll_disconnect.etl");

        // Start capture in background
        RunCommand("logman", "stop deeppoll -ets", silent: true);
        RunCommand("logman", $"start deeppoll -p Microsoft-Windows-USB-UCX -o \"{etlPath}\" -ets");

        Console.WriteLine();
        Console.WriteLine("  Logging... Press ENTER when disconnect happens.");
        Console.WriteLine();

        // Wait for user or timeout (30 min)
        var startTime = DateTime.Now;
        while (!Console.KeyAvailable && (DateTime.Now - startTime).TotalMinutes < 30)
        {
            var elapsed = DateTime.Now - startTime;
            Console.Write($"\r  Recording: {elapsed:mm\\:ss}  ");
            Thread.Sleep(1000);
        }
        Console.ReadKey(true); // consume the key

        // Stop capture
        RunCommand("logman", "stop deeppoll -ets");

        Console.WriteLine();
        Console.WriteLine();
        Console.WriteLine("  Processing...");
        Console.WriteLine();

        AnalyzeEtl(etlPath, false, "054C:05C4");

        string? logId = UploadLog(etlPath, "disconnect");
        if (logId != null)
        {
            Console.WriteLine();
            Console.WriteLine($"  ┌─────────────────────────────────────────────────────┐");
            Console.WriteLine($"  │  Upload complete!                                    │");
            Console.WriteLine($"  │                                                     │");
            Console.WriteLine($"  │  Log ID: {logId,-43} │");
            Console.WriteLine($"  │                                                     │");
            Console.WriteLine($"  │  Send this ID to Marius on Discord:                 │");
            Console.WriteLine($"  │  https://discord.gg/4Q9SRUt85j                      │");
            Console.WriteLine($"  └─────────────────────────────────────────────────────┘");
        }

        try { File.Delete(etlPath); } catch { }
    }

    static void StartEtwCapture(string etlPath, int seconds)
    {
        // Stop any existing trace
        RunCommand("logman", "stop usbpollcap -ets", silent: true);

        // Start new trace
        RunCommand("logman", $"start usbpollcap -p Microsoft-Windows-USB-UCX -o \"{etlPath}\" -ets");

        // Wait
        Thread.Sleep(seconds * 1000);

        // Stop trace
        RunCommand("logman", "stop usbpollcap -ets");
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

    static void AnalyzeEtl(string etlPath, bool verbose, string? filterDevice = null)
    {
        if (!File.Exists(etlPath))
        {
            Console.WriteLine("  Capture failed - no data collected.");
            return;
        }

        // Parse ETL
        var (interruptTx, controlTx, captureDurationMs) = ParseEtl(etlPath);

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

                // Estimate poll rate from average interval
                double avgIntervalUs = 0;
                if (txList.Count > 1)
                {
                    var intervals = new List<double>();
                    for (int i = 1; i < txList.Count; i++)
                    {
                        double interval = (txList[i].EndTimestamp - txList[i - 1].EndTimestamp) * 1000;
                        if (interval > 0 && interval < 50000) intervals.Add(interval);
                    }
                    if (intervals.Count > 0) avgIntervalUs = intervals.Average();
                }
                double estimatedHz = avgIntervalUs > 0 ? 1000000.0 / avgIntervalUs : 0;

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

        List<UsbTransaction> selectedTransactions;

        // If filterDevice specified, try to use that
        if (!string.IsNullOrEmpty(filterDevice))
        {
            var filtered = deviceGroups.FirstOrDefault(g => g.VidPid == filterDevice);
            if (filtered != null)
            {
                selectedTransactions = filtered.Transactions;
            }
            else
            {
                // VID:PID not in ETL data - auto-select device with most samples
                selectedTransactions = deviceGroups[0].Transactions;
            }
        }
        // If only one device, auto-select
        else if (deviceGroups.Count == 1)
        {
            selectedTransactions = deviceGroups[0].Transactions;
        }
        // Multiple devices - let user select
        else
        {
            Console.WriteLine();
            Console.WriteLine("  Found devices:");
            Console.WriteLine();
            for (int i = 0; i < deviceGroups.Count; i++)
            {
                var g = deviceGroups[i];
                string hzStr = g.EstimatedHz > 0 ? $"~{g.EstimatedHz:F0} Hz" : "";
                Console.WriteLine($"    [{i + 1}] {g.VidPid,-12} {g.Count,6:N0} samples  {hzStr}");
            }
            Console.WriteLine();
            Console.Write("  Select: ");

            string? choice = Console.ReadLine()?.Trim();
            if (!int.TryParse(choice, out int idx) || idx < 1 || idx > deviceGroups.Count)
            {
                Console.WriteLine("  Invalid selection.");
                return;
            }

            selectedTransactions = deviceGroups[idx - 1].Transactions;
        }

        // Calculate intervals for selected device
        var sorted = selectedTransactions;
        var intervals = new List<double>();
        for (int i = 1; i < sorted.Count; i++)
        {
            double interval = (sorted[i].EndTimestamp - sorted[i - 1].EndTimestamp) * 1000;
            if (interval > 0 && interval < 50000) intervals.Add(interval);
        }

        if (intervals.Count == 0)
        {
            Console.WriteLine("  Not enough data to calculate poll rate.");
            return;
        }

        var sortedIntervals = intervals.OrderBy(x => x).ToList();

        // Calculate stats
        double pollRate = 1000000.0 / intervals.Average();
        double p50 = Percentile(sortedIntervals, 50);
        double p90 = Percentile(sortedIntervals, 90);
        double p99 = Percentile(sortedIntervals, 99);
        double p999 = Percentile(sortedIntervals, 99.9);
        double maxInterval = sortedIntervals.Last();

        int gaps200 = intervals.Count(x => x > 200);
        int gaps500 = intervals.Count(x => x > 500);
        int gaps1ms = intervals.Count(x => x > 1000);
        int gaps5ms = intervals.Count(x => x > 5000);

        // Pretty output
        Console.WriteLine();
        PrintDoubleLine(60);
        PrintCentered("D E E P P O L L", 60);
        PrintDoubleLine(60);

        Console.WriteLine($"  Poll Rate:  {pollRate:F0} Hz    Samples: {intervals.Count:N0}");
        Console.WriteLine();
        Console.WriteLine("  Timing Distribution:");
        Console.WriteLine();

        // Auto-bin based on data range
        double medianUs = Percentile(sortedIntervals, 50);
        double p5 = Percentile(sortedIntervals, 5);
        double p95 = Percentile(sortedIntervals, 95);

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
        }
    }

    static (List<UsbTransaction> interrupt, List<UsbTransaction> control, double durationMs) ParseEtl(string etlPath)
    {
        var pendingTransactions = new Dictionary<ulong, UsbTransaction>();
        var completedTransactions = new List<UsbTransaction>();
        double firstTimestamp = -1;
        double lastTimestamp = 0;

        using (var source = new ETWTraceEventSource(etlPath))
        {
            source.Dynamic.All += (TraceEvent data) =>
            {
                string provider = data.ProviderName ?? "";
                string eventName = data.EventName ?? "";

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

        var interrupt = completedTransactions.Where(t => t.Type == "Interrupt").ToList();
        var control = completedTransactions.Where(t => t.Type == "Control").ToList();

        return (interrupt, control, lastTimestamp - firstTimestamp);
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

    static string? UploadLog(string etlPath, string logType)
    {
        Console.WriteLine();
        Console.WriteLine("  ┌─────────────────────────────────────────────────────┐");
        Console.WriteLine("  │  Upload log to Marius for analysis?                 │");
        Console.WriteLine("  └─────────────────────────────────────────────────────┘");
        Console.WriteLine();
        Console.WriteLine("    [1] Yes, upload");
        Console.WriteLine("    [2] No, skip");
        Console.WriteLine();
        Console.Write("  Select: ");

        string? choice = Console.ReadLine()?.Trim();
        if (choice != "1") return null;

        Console.WriteLine();
        Console.Write("  Your nickname (Discord/name): ");
        string? nickname = Console.ReadLine()?.Trim();
        if (string.IsNullOrEmpty(nickname))
        {
            Console.WriteLine("  Skipped - no nickname provided.");
            return null;
        }

        Console.WriteLine();
        Console.WriteLine("  Compressing...");

        // Compress the ETL file
        string gzPath = etlPath + ".gz";
        try
        {
            using (var input = File.OpenRead(etlPath))
            using (var output = File.Create(gzPath))
            using (var gz = new GZipStream(output, CompressionLevel.Optimal))
            {
                input.CopyTo(gz);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Compression failed: {ex.Message}");
            return null;
        }

        long fileSize = new FileInfo(gzPath).Length;
        Console.WriteLine($"  Compressed: {fileSize / 1024 / 1024} MB");
        Console.WriteLine();
        Console.WriteLine("  Uploading...");

        try
        {
            using var client = new HttpClient();
            client.Timeout = TimeSpan.FromMinutes(10);

            // Get presigned URL
            var requestBody = JsonSerializer.Serialize(new
            {
                nickname = nickname,
                logType = logType,
                fileSize = fileSize
            });

            var urlResponse = client.PostAsync(
                "https://tools.mariusheier.com/deeppoll/upload-url",
                new StringContent(requestBody, System.Text.Encoding.UTF8, "application/json")
            ).Result;

            if (!urlResponse.IsSuccessStatusCode)
            {
                Console.WriteLine($"  Failed to get upload URL: {urlResponse.StatusCode}");
                try { File.Delete(gzPath); } catch { }
                return null;
            }

            var urlData = JsonSerializer.Deserialize<JsonElement>(urlResponse.Content.ReadAsStringAsync().Result);
            string uploadUrl = urlData.GetProperty("url").GetString() ?? "";
            string logId = urlData.GetProperty("id").GetString() ?? "";

            // Upload file
            using var fileContent = new ByteArrayContent(File.ReadAllBytes(gzPath));
            fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/gzip");

            var uploadResponse = client.PutAsync(uploadUrl, fileContent).Result;

            try { File.Delete(gzPath); } catch { }

            if (!uploadResponse.IsSuccessStatusCode)
            {
                Console.WriteLine($"  Upload failed: {uploadResponse.StatusCode}");
                return null;
            }

            return logId;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  Upload failed: {ex.Message}");
            try { File.Delete(gzPath); } catch { }
            return null;
        }
    }
}
