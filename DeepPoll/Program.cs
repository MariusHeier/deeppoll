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

    public double DurationUs => (EndTimestamp - StartTimestamp) * 1000.0;
    public bool IsComplete => EndTimestamp > 0;
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
            // Other USB device - ask what type
            Console.WriteLine();
            PrintSingleLine(60);
            Console.WriteLine();
            Console.WriteLine("  What type of device?");
            Console.WriteLine();
            Console.WriteLine("    [1] Mouse");
            Console.WriteLine("    [2] Gamepad / Controller");
            Console.WriteLine("    [3] Keyboard");
            Console.WriteLine("    [4] Other");
            Console.WriteLine();
            Console.Write("  Select: ");

            string? deviceType = Console.ReadLine()?.Trim();
            string deviceInstruction = deviceType switch
            {
                "1" => "Keep MOVING your mouse during the capture!",
                "2" => "Keep PRESSING buttons or moving sticks!",
                "3" => "Keep PRESSING keys during the capture!",
                _ => "Keep using the device during capture!"
            };

            RunPollCheck(verbose, null, deviceInstruction);
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
            RunPollCheck(verbose, "054C:05C4", "Keep PRESSING buttons or moving sticks!");
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

    static void RunPollCheck(bool verbose, string? deviceFilter, string deviceInstruction = "Keep using the device during capture!")
    {
        Console.WriteLine();
        Console.WriteLine("  ┌─────────────────────────────────────────────────────┐");
        Console.WriteLine("  │  Make sure your device is plugged in.               │");
        Console.WriteLine("  │                                                     │");
        Console.WriteLine($"  │  {deviceInstruction,-51} │");
        Console.WriteLine("  │                                                     │");
        Console.WriteLine("  │  Press ENTER to start 7-second capture.             │");
        Console.WriteLine("  └─────────────────────────────────────────────────────┘");
        Console.ReadLine();

        string etlPath = Path.Combine(Path.GetTempPath(), "deeppoll_capture.etl");

        Console.WriteLine();
        Console.WriteLine("  Capturing for 7 seconds...");
        Console.WriteLine();

        StartEtwCapture(etlPath, 7);

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

        // TODO: Compress and upload
        Console.WriteLine();
        Console.WriteLine("  [Upload not implemented yet]");

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

        // TODO: Compress and upload
        Console.WriteLine();
        Console.WriteLine("  [Upload not implemented yet]");

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
            Console.WriteLine("No interrupt transfers found in trace.");
            return;
        }

        // Calculate intervals
        var sorted = interruptTx.OrderBy(t => t.EndTimestamp).ToList();
        var intervals = new List<double>();
        for (int i = 1; i < sorted.Count; i++)
        {
            double interval = (sorted[i].EndTimestamp - sorted[i - 1].EndTimestamp) * 1000;
            if (interval > 0 && interval < 50000) intervals.Add(interval);
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

        // Bin by Hz (convert from interval µs)
        // interval µs → Hz = 1,000,000 / interval
        var bins = new (string Label, double MinUs, double MaxUs)[]
        {
            ("8000+ Hz", 0, 125),        // < 125µs = > 8000 Hz
            ("7700-8000", 125, 130),     // 125-130µs
            ("7200-7700", 130, 139),     // 130-139µs
            ("5000-7200", 139, 200),     // 139-200µs
            ("<5000 Hz", 200, double.MaxValue)  // > 200µs = < 5000 Hz
        };

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

            char barChar = i == bins.Length - 1 ? '░' : '▓';
            string bar = new string(barChar, barLen);

            Console.WriteLine($"  {bins[i].Label,10} {bar.PadRight(35)} {pct,5:F1}%");
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
                    try { deviceHandle = Convert.ToUInt64(data.PayloadByName("fid_UsbDevice")); } catch { }
                    try { pipeHandle = Convert.ToUInt64(data.PayloadByName("fid_PipeHandle")); } catch { }

                    pendingTransactions[urbPtr] = new UsbTransaction
                    {
                        UrbPointer = urbPtr,
                        DeviceHandle = deviceHandle,
                        PipeHandle = pipeHandle,
                        Type = txType,
                        StartTimestamp = data.TimeStampRelativeMSec
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
}
