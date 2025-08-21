using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;

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

        // ---------- Tipagem: Janela/Tela ----------
        private static string InferTipo(string displayPath, IList<(IntPtr h, string cls, string title, (int X, int Y, int W, int H) rc)> childList)
        {
            // 1) Convenção de nomes (mais robusta no FTView do seu projeto)
            var nome = DeriveTokenFromDisplay(displayPath);
            if (nome.StartsWith("jan_", StringComparison.OrdinalIgnoreCase)) return "Janela";
            if (nome.StartsWith("tela_", StringComparison.OrdinalIgnoreCase)) return "Tela";
            if (nome.Equals("cabecalho", StringComparison.OrdinalIgnoreCase) || nome.StartsWith("cabecalho_", StringComparison.OrdinalIgnoreCase))
                return "Janela"; // costuma vir com menus/ações
            if (nome.Equals("rodape", StringComparison.OrdinalIgnoreCase) || nome.StartsWith("rodape", StringComparison.OrdinalIgnoreCase))
                return "Janela"; // mesma lógica

            // 2) Heurística visual: presença de barras/menus/MDI child com controles "de barra"
            // Procura classes comuns de barra/menu/toolbar (não exaustivo; ajustável)
            string[] barraClasses = { "ReBarWindow32", "ToolbarWindow32", "MsoCommandBarDock", "MsoCommandBar", "ATL:", "Afx:ControlBar", "AfxMDIFrame" };
            bool temBarra = childList.Any(c => barraClasses.Any(b => c.cls.IndexOf(b, StringComparison.OrdinalIgnoreCase) >= 0));
            if (temBarra) return "Janela";

            // 3) Fallback
            return "Tela";
        }

        // ---------- Snapshot FTView ----------
        /// <summary>
        /// Faz snapshot do container FTView (DisplayClient/ViewSE/SEClient), lista alguns filhos e infere 'Tipo=Janela|Tela'.
        /// </summary>
        public static void SnapshotFtView(string displayPath, Action<string> log, int childLimit = 8)
        {
            string[] procNames = { "DisplayClient", "ViewSE", "SEClient", "FTViewSEClient" };
            foreach (var n in procNames)
            {
                foreach (var p in Process.GetProcessesByName(n))
                {
                    // tenta janela principal; se não houver, varre top-level por PID
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

                        // coleta filhos
                        var children = new List<(IntPtr h, string cls, string title, (int, int, int, int) rc)>();
                        int count = 0;
                        foreach (var ch in EnumChildren(main))
                        {
                            var cTitle = GetWindowTitle(ch);
                            var cClass = GetWindowClass(ch);
                            var rc = GetBounds(ch);
                            children.Add((ch, cClass, cTitle, rc));
                            if (count++ < childLimit)
                                log?.Invoke($"[HWND][FTView][Child] handle=0x{ch.ToInt64():X} class='{cClass}' title='{cTitle}' pos=({rc.X},{rc.Y}) size=({rc.W}x{rc.H})");
                        }

                        // classificar abertura
                        var tipo = InferTipo(displayPath, children);
                        log?.Invoke($"[HWND][FTView][Tipo] display='{displayPath}' => '{tipo}'");
                        return; // 1° container suficiente
                    }
                }
            }
            log?.Invoke($"[HWND][FTView] Nenhum container FTView encontrado (processo).");
        }
    }
}
