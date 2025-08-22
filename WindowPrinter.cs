// LibFTView\Win32\WindowPrinter.cs
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;

namespace LibFTView.Win32
{
    internal static class WindowPrinter
    {
        private const int SRCCOPY = 0x00CC0020;
        private const int PW_RENDERFULLCONTENT = 0x00000002;

        [DllImport("user32.dll")] private static extern bool PrintWindow(IntPtr hwnd, IntPtr hdcBlt, int nFlags);
        [DllImport("user32.dll")] private static extern IntPtr GetDC(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr hdc);
        [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr hdc);
        [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);
        [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hgdiobj);
        [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr hObject);
        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdc, int x, int y, int cx, int cy,
                                                                    IntPtr hdcSrc, int x1, int y1, int rop);

        public static bool CaptureToPng(IntPtr hwnd, int x, int y, int w, int h, string filePath,
                                        Action<string> log, out string method)
        {
            method = "BitBlt";
            if (hwnd == IntPtr.Zero || w <= 0 || h <= 0)
            {
                log?.Invoke("[Print][Skip] hwnd inválido ou tamanho zero.");
                return false;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath)!);

            var sw = Stopwatch.StartNew();
            IntPtr hdcScreen = IntPtr.Zero, hdcMem = IntPtr.Zero, hBmp = IntPtr.Zero, hOld = IntPtr.Zero;
            try
            {
                hdcScreen = GetDC(IntPtr.Zero);
                hdcMem = CreateCompatibleDC(hdcScreen);
                hBmp = CreateCompatibleBitmap(hdcScreen, w, h);
                hOld = SelectObject(hdcMem, hBmp);

                if (PrintWindow(hwnd, hdcMem, PW_RENDERFULLCONTENT))
                    method = "PrintWindow";
                else
                    BitBlt(hdcMem, 0, 0, w, h, hdcScreen, x, y, SRCCOPY);

                using var bmp = Image.FromHbitmap(hBmp);
                bmp.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);

                sw.Stop();
                var fi = new FileInfo(filePath);
                log?.Invoke($"[Print][Saved] file='{filePath}' size=({w}x{h}) bytes={fi.Length} method='{method}' elapsedMs={sw.ElapsedMilliseconds}");
                return true;
            }
            catch (Exception ex)
            {
                log?.Invoke($"[Print][Error] {ex.GetType().Name}: {ex.Message}");
                return false;
            }
            finally
            {
                if (hOld != IntPtr.Zero) SelectObject(hdcMem, hOld);
                if (hBmp != IntPtr.Zero) DeleteObject(hBmp);
                if (hdcMem != IntPtr.Zero) DeleteDC(hdcMem);
                if (hdcScreen != IntPtr.Zero) ReleaseDC(IntPtr.Zero, hdcScreen);
            }
        }
    }
}
