using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using System.Drawing;            // <-- ADICIONE
using System.Drawing.Imaging;    // <-- ADICIONE
using System.IO;                 // <-- ADICIONE
using System.Threading;          // <-- ADICIONE

namespace LibFTView.Win32
{
    internal static class WindowProbe
    {
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        // ------ CONFIG PRINT ------
        private static readonly string PrintsDir =
            @"C:\RSLogix 5000\Projects\P83\Telas\TELAS_XML\GFX\Prints";

        private static void EnsurePrintsDir()
        {
            try { Directory.CreateDirectory(PrintsDir); } catch { /* ignore */ }
        }

        private static string SanitizeFileName(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "display";
            foreach (var ch in Path.GetInvalidFileNameChars()) s = s.Replace(ch, '_');
            return s;
        }

        // ------ HELPERS WINDOW ------
        private static string GetWindowTitle(IntPtr h)
        {
            int len = GetWindowTextLength(h);
            if (len <= 0) return string.Empty;
            var sb = new StringBuilder(len + 1);
            GetWindowText(h, sb, sb.Capacity);
            return sb.ToString().Trim();
        }

        private static string GetWindowClass(IntPtr h)
        {
            var sb = new StringBuilder(256);
            GetClassName(h, sb, sb.Capacity);
            return sb.ToString();
        }

        private static (int X, int Y, int W, int H) GetBounds(IntPtr h)
        {
            if (GetWindowRect(h, out var r))
                return (r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
            return (0, 0, 0, 0);
        }

        private static IEnumerable<IntPtr> EnumTopLevelWindowsByPid(int pid)
        {
            var list = new List<IntPtr>();
            EnumWindows((h, l) =>
            {
                if (!IsWindowVisible(h)) return true;
                GetWindowThreadProcessId(h, out var wp);
                if ((int)wp == pid) list.Add(h);
                return true;
            }, IntPtr.Zero);
            return list;
        }

        private static IEnumerable<IntPtr> EnumChildren(IntPtr parent)
        {
            var list = new List<IntPtr>();
            if (parent == IntPtr.Zero) return list;
            EnumChildWindows(parent, (h, l) => { list.Add(h); return true; }, IntPtr.Zero);
            return list;
        }

        private static string SafeExePathFromPid(uint pid)
        {
            try { return Process.GetProcessById((int)pid)?.MainModule?.FileName ?? ""; }
            catch { return ""; }
        }

        private static string DeriveTokenFromDisplay(string displayPath)
        {
            if (string.IsNullOrWhiteSpace(displayPath)) return "";
            var s = displayPath.Trim();
            var idx = s.IndexOf("::", StringComparison.Ordinal);
            if (idx >= 0 && idx + 2 < s.Length) s = s.Substring(idx + 2);
            return s; // ex.: "jan_ger_..." ou "l1_..."
        }

        private static bool WaitForStableBounds(IntPtr h, int timeoutMs = 1000, int polls = 5, int tol = 1)
        {
            var prev = GetBounds(h);
            int stableCount = 0;
            for (int i = 0; i < polls; i++)
            {
                Thread.Sleep(Math.Max(1, timeoutMs / polls));
                var cur = GetBounds(h);
                bool eq = Math.Abs(cur.X - prev.X) <= tol &&
                          Math.Abs(cur.Y - prev.Y) <= tol &&
                          Math.Abs(cur.W - prev.W) <= tol &&
                          Math.Abs(cur.H - prev.H) <= tol;
                if (eq) stableCount++;
                else stableCount = 0;

                prev = cur;
                if (stableCount >= 2) return true; // ~estável em 2 leituras seguidas
            }
            return false; // segue imprimindo mesmo assim
        }

        private static bool RectValid((int X, int Y, int W, int H) rc) => rc.W > 2 && rc.H > 2;

        // Adicione dentro da classe WindowProbe
        public static bool TryGetFtViewRenderTarget(out IntPtr renderHwnd, out string fullFromTitle, Action<string> log)
        {
            renderHwnd = IntPtr.Zero;
            fullFromTitle = string.Empty;

            string[] procNames = { "DisplayClient", "ViewSE", "SEClient", "FTViewSEClient" };

            foreach (var n in procNames)
            {
                foreach (var p in Process.GetProcessesByName(n))
                {
                    // preferir MainWindow; se não, procurar top-levels do PID
                    var tops = new List<IntPtr>();
                    if (p.MainWindowHandle != IntPtr.Zero) tops.Add(p.MainWindowHandle);
                    else tops.AddRange(EnumTopLevelWindowsByPid(p.Id));

                    foreach (var main in tops)
                    {
                        // enumerar filhos e escolher o candidato com título de display
                        var children = new List<(IntPtr h, string cls, string title, (int X, int Y, int W, int H) rc)>();
                        foreach (var ch in EnumChildren(main))
                        {
                            var ttl = GetWindowTitle(ch);
                            var cls = GetWindowClass(ch);
                            var rc = GetBounds(ch);
                            children.Add((ch, cls, ttl, rc));
                        }

                        // critério 1: título que contém " - /TELAS" e maior área
                        var candidates = children
                            .Where(c => !string.IsNullOrWhiteSpace(c.title) &&
                                        c.title.IndexOf(" - /TELAS", StringComparison.OrdinalIgnoreCase) >= 0 &&
                                        c.rc.W > 0 && c.rc.H > 0)
                            .OrderByDescending(c => (long)c.rc.W * c.rc.H)
                            .ToList();

                        var target = candidates.FirstOrDefault();
                        if (target.h == IntPtr.Zero)
                        {
                            // critério 2 (fallback): qualquer filho grande sob MDIClient
                            var mdi = children.FirstOrDefault(c => c.cls.Equals("MDIClient", StringComparison.OrdinalIgnoreCase)).h;
                            var underMdi = children
                                .Where(c => c.h != mdi && c.rc.W > 300 && c.rc.H > 200)
                                .OrderByDescending(c => (long)c.rc.W * c.rc.H)
                                .FirstOrDefault();
                            target = underMdi;
                        }

                        if (target.h != IntPtr.Zero)
                        {
                            // Montar "/AREA::NOME" a partir do título "nome - /TELAS_...//AREA"
                            string area = "";
                            string nome = target.title;
                            var dash = nome.IndexOf(" - ", StringComparison.Ordinal);
                            if (dash > 0) nome = nome.Substring(0, dash);

                            var idxArea = target.title.LastIndexOf("//", StringComparison.Ordinal);
                            if (idxArea >= 0 && idxArea + 2 < target.title.Length)
                                area = target.title.Substring(idxArea + 2).Trim();

                            if (!string.IsNullOrWhiteSpace(area) && !string.IsNullOrWhiteSpace(nome))
                                fullFromTitle = "/" + area + "::" + nome;

                            renderHwnd = target.h;
                            log?.Invoke($"[Print][Target] hwnd=0x{renderHwnd.ToInt64():X} class='{target.cls}' title='{target.title}' pos=({target.rc.X},{target.rc.Y}) size=({target.rc.W}x{target.rc.H}) full='{fullFromTitle}'");
                            return true;
                        }
                    }
                }
            }

            log?.Invoke("[Print][Target] Não encontrado (DisplayClient/ViewSE).");
            return false;
        }


        // ---------- Tipagem: Janela/Tela ----------
        private static string InferTipo(string displayPath, IList<(IntPtr h, string cls, string title, (int X, int Y, int W, int H) rc)> childList)
        {
            var nome = DeriveTokenFromDisplay(displayPath);
            if (nome.StartsWith("jan_", StringComparison.OrdinalIgnoreCase)) return "Janela";
            if (nome.StartsWith("tela_", StringComparison.OrdinalIgnoreCase)) return "Tela";
            if (nome.Equals("cabecalho", StringComparison.OrdinalIgnoreCase) || nome.StartsWith("cabecalho_", StringComparison.OrdinalIgnoreCase))
                return "Janela";
            if (nome.Equals("rodape", StringComparison.OrdinalIgnoreCase) || nome.StartsWith("rodape", StringComparison.OrdinalIgnoreCase))
                return "Janela";

            string[] barraClasses = { "ReBarWindow32", "ToolbarWindow32", "MsoCommandBarDock", "MsoCommandBar", "ATL:", "Afx:ControlBar", "AfxMDIFrame" };
            bool temBarra = childList.Any(c => barraClasses.Any(b => c.cls.IndexOf(b, StringComparison.OrdinalIgnoreCase) >= 0));
            if (temBarra) return "Janela";

            return "Tela";
        }

        // ---------- PRINT ----------
        private static void TryPrintSelectedChild(string displayPath, IList<(IntPtr h, string cls, string title, (int X, int Y, int W, int H) rc)> children, Action<string> log)
        {
            try
            {
                var token = DeriveTokenFromDisplay(displayPath);
                var safeName = SanitizeFileName(token);
                EnsurePrintsDir();

                // 1) escolhe o filho pelo título que contém o token + " - "
                var target = children.FirstOrDefault(c =>
                    !string.IsNullOrWhiteSpace(c.title) &&
                    (c.title.StartsWith(token, StringComparison.OrdinalIgnoreCase) ||
                     c.title.IndexOf(token + " -", StringComparison.OrdinalIgnoreCase) >= 0));

                if (target.h == IntPtr.Zero)
                {
                    log?.Invoke($"[Print][Skip] display='{displayPath}' token='{token}': nenhum filho com título correspondente.");
                    return;
                }

                log?.Invoke($"[Print][Pick] display='{displayPath}' token='{token}' handle=0x{target.h.ToInt64():X} class='{target.cls}' title='{target.title}' rect=({target.rc.X},{target.rc.Y},{target.rc.W}x{target.rc.H})");

                // 2) espera estabilizar o retângulo (leve debounce)
                bool stable = WaitForStableBounds(target.h, timeoutMs: 1000, polls: 5, tol: 1);
                log?.Invoke($"[Print][Wait] estabilizado={(stable ? "sim" : "não")} timeout=1000ms");

                var rc = GetBounds(target.h); // pega bounds atualizados pós-wait
                if (!RectValid(rc))
                {
                    log?.Invoke($"[Print][ERR] retângulo inválido ({rc.X},{rc.Y},{rc.W}x{rc.H}). Abortando.");
                    return;
                }

                // 3) caminho final (evita overwrite)
                string baseFile = Path.Combine(PrintsDir, $"{safeName}.png");
                string finalFile = baseFile;
                if (File.Exists(finalFile))
                    finalFile = Path.Combine(PrintsDir, $"{safeName}_{DateTime.Now:yyyyMMdd_HHmmss}.png");

                // 4) captura por tela (funciona bem com MDI/Afx do FTView)
                using (var bmp = new Bitmap(rc.W, rc.H, PixelFormat.Format32bppArgb))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(rc.X, rc.Y, 0, 0, new Size(rc.W, rc.H));
                    bmp.Save(finalFile, ImageFormat.Png);
                }

                log?.Invoke($"[Print][OK] name='{safeName}.png' size=({rc.W}x{rc.H}) saved='{finalFile}'");
            }
            catch (Exception ex)
            {
                log?.Invoke($"[Print][ERR] exception='{ex.GetType().Name}: {ex.Message}'");
            }
        }

        // ---------- Snapshot FTView ----------
        /// <summary>
        /// Faz snapshot do container FTView (DisplayClient/ViewSE/SEClient), lista alguns filhos,
        /// infere 'Tipo=Janela|Tela' e DISPARA PRINT do filho-alvo (título contém o token).
        /// </summary>
        public static void SnapshotFtView(string displayPath, Action<string> log, int childLimit = 8)
        {
            string[] procNames = { "DisplayClient", "ViewSE", "SEClient", "FTViewSEClient" };
            foreach (var n in procNames)
            {
                foreach (var p in Process.GetProcessesByName(n))
                {
                    var tops = new List<IntPtr>();
                    if (p.MainWindowHandle != IntPtr.Zero) tops.Add(p.MainWindowHandle);
                    else tops.AddRange(EnumTopLevelWindowsByPid(p.Id));

                    foreach (var main in tops)
                    {
                        var title = GetWindowTitle(main);
                        var cls = GetWindowClass(main);
                        var (x, y, w, h) = GetBounds(main);
                        var exe = SafeExePathFromPid((uint)p.Id);

                        log?.Invoke($"[HWND][FTView][Container] handle=0x{main.ToInt64():X} pid={p.Id} exe='{exe}' class='{cls}' title='{title}' pos=({x},{y}) size=({w}x{h})");

                        var children = new List<(IntPtr h, string cls, string title, (int, int, int, int) rc)>();
                        int count = 0;
                        foreach (var ch in EnumChildren(main))
                        {
                            var cTitle = GetWindowTitle(ch);
                            var cClass = GetWindowClass(ch);
                            var rc = GetBounds(ch);
                            children.Add((ch, cClass, cTitle, rc));
                            if (count++ < childLimit)
                                log?.Invoke($"[HWND][FTView][Child] handle=0x{ch.ToInt64():X} class='{cClass}' title='{cTitle}' pos=({rc.Item1},{rc.Item2}) size=({rc.Item3}x{rc.Item4})");
                        }

                        var tipo = InferTipo(displayPath, children);
                        log?.Invoke($"[HWND][FTView][Tipo] display='{displayPath}' => '{tipo}'");

                        // --------- INJEÇÃO DO PRINT (AQUI) ---------
                        TryPrintSelectedChild(displayPath, children, log);
                        // -------------------------------------------

                        return; // 1° container suficiente
                    }
                }
            }
            log?.Invoke($"[HWND][FTView] Nenhum container FTView encontrado (processo).");
        }
    }
}
