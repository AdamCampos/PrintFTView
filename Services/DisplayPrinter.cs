using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace LibFTView.Services
{
    public static class DisplayPrinter
    {
        public static string DefaultOutDir = @"C:\RSLogix 5000\Projects\P83\Telas\TELAS_XML\GFX\Prints";

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        private static (int X, int Y, int W, int H) GetBounds(IntPtr h)
        {
            if (GetWindowRect(h, out var r)) return (r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
            return (0, 0, 0, 0);
        }

        private static (int X, int Y, int W, int H) WaitStableRect(IntPtr h, int maxWaitMs, int sampleIntervalMs, int requiredStableSamples)
        {
            var last = GetBounds(h);
            int stable = 0;
            var started = Environment.TickCount;

            while (Environment.TickCount - started <= maxWaitMs)
            {
                Thread.Sleep(sampleIntervalMs);
                var cur = GetBounds(h);
                if (cur.Equals(last))
                {
                    if (++stable >= requiredStableSamples) return cur;
                }
                else
                {
                    stable = 0;
                    last = cur;
                }
            }
            return last;
        }

        private static string TokenFromDisplayOrTitle(string full)
        {
            if (string.IsNullOrWhiteSpace(full)) return "display";
            var pos = full.IndexOf("::", StringComparison.Ordinal);
            if (pos >= 0 && pos + 2 < full.Length)
                full = full.Substring(pos + 2);
            var dash = full.IndexOf(" - ", StringComparison.Ordinal);
            if (dash > 0) full = full.Substring(0, dash);
            foreach (var c in Path.GetInvalidFileNameChars()) full = full.Replace(c, '_');
            return full.Trim();
        }

        public static string CaptureAfterStable(
            IntPtr hwnd,
            string fullDisplayOrTitle,
            string outDir,
            int maxWaitMs,
            int sampleIntervalMs,
            int requiredStableSamples,
            Action<string> log)
        {
            try { Directory.CreateDirectory(outDir); } catch { }

            var rect = WaitStableRect(hwnd, maxWaitMs, sampleIntervalMs, requiredStableSamples);
            if (rect.W <= 0 || rect.H <= 0)
            {
                log?.Invoke($"[Print][Skip] retângulo inválido (W/H <= 0) hwnd=0x{hwnd.ToInt64():X}");
                return string.Empty;
            }

            var token = TokenFromDisplayOrTitle(fullDisplayOrTitle);
            var outPath = Path.Combine(outDir, token + ".png");

            log?.Invoke($"[Print][Start] display='{fullDisplayOrTitle}' hwnd=0x{hwnd.ToInt64():X} rect=({rect.X},{rect.Y}) size=({rect.W}x{rect.H}) -> '{outPath}'");

            if (LibFTView.Win32.WindowPrinter.CaptureToPng(hwnd, rect.X, rect.Y, rect.W, rect.H, outPath, log, out var method))
            {
                log?.Invoke($"[Print][Done] file='{outPath}' size=({rect.W}x{rect.H}) method='{method}'");
                return outPath;
            }

            log?.Invoke("[Print][Fail] Não foi possível capturar.");
            return string.Empty;
        }
    }
}
