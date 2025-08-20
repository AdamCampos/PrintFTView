// FTViewCom.cs — versão recriada/simplificada
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace LibFTView.Services
{
    [ComVisible(true)]
    [Guid("6D6F9F2E-7E77-4F7A-9B47-2B35B8B8E9D9")]
    [ProgId("LibFTView.FTViewCom")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class FTViewCom
    {
        // ===== Config de log/diagnóstico =====
        // Default exigido pelo usuário:
        private string _diagLogPathHwnd = @"C:\Projetos\VisualStudio\LibFTView\ftviewcom_hwnd.log";
        private bool _diagShowMsgBoxHwnd = true;

        public FTViewCom()
        {
            try
            {
                var dir = Path.GetDirectoryName(_diagLogPathHwnd);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir!);
            }
            catch { /* não interromper */ }
        }

        /// <summary>
        /// Ajusta exibição de MessageBox e o caminho do log (cria pasta se necessário).
        /// </summary>
        [DispId(95)]
        public void SetHwndDiagOptions(bool showMessageBoxes, string? logPath)
        {
            _diagShowMsgBoxHwnd = showMessageBoxes;
            if (!string.IsNullOrWhiteSpace(logPath))
                _diagLogPathHwnd = logPath!;
            try
            {
                var dir = Path.GetDirectoryName(_diagLogPathHwnd);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir!);
            }
            catch { /* ignore */ }
        }

        /// <summary>
        /// Atalho: define apenas o diretório (arquivo "ftviewcom_hwnd.log").
        /// </summary>
        [DispId(95 + 1000)]
        public void SetDiagRoot(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder)) return;
            try { Directory.CreateDirectory(folder); } catch { }
            _diagLogPathHwnd = Path.Combine(folder, "ftviewcom_hwnd.log");
        }

        /// <summary>
        /// Processa a tela já aberta (pelo Display do FTView/VBA), encontra a janela de exibição,
        /// tira screenshot da ÁREA CLIENTE e registra log. Retorna "OK: ..." ou "#ERR: ...".
        /// </summary>
        /// <param name="expectedPath">Ex.: "AREA\TELA" (apenas para mostrar no MsgBox/log)</param>
        /// <param name="outPngFullPath">Caminho completo para salvar o PNG</param>
        /// <param name="waitMs">Tempo (ms) para a janela estabilizar</param>
        [DispId(96)]
        public string ProcessDisplayByHwnd(string expectedPath, string outPngFullPath, int waitMs = 250)
        {
            void Log(string s)
            {
                try
                {
                    var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}  {s}";
                    File.AppendAllText(_diagLogPathHwnd, line + Environment.NewLine, Encoding.UTF8);
                }
                catch { /* ignore */ }
            }
            void Mb(string s)
            {
                if (!_diagShowMsgBoxHwnd) return;
                try { MessageBox.Show(s, "FTViewCom/HWND"); } catch { }
            }

            try
            {
                if (string.IsNullOrWhiteSpace(outPngFullPath))
                    throw new ArgumentException("outPngFullPath vazio");

                var d = Path.GetDirectoryName(outPngFullPath);
                if (!string.IsNullOrEmpty(d) && !Directory.Exists(d)) Directory.CreateDirectory(d);

                Log("=== ProcessDisplayByHwnd START ===");
                Log($"expectedPath='{expectedPath}' | out='{outPngFullPath}' | wait={waitMs}ms");

                if (waitMs > 0)
                    System.Threading.Thread.Sleep(waitMs);

                // 1) Acha a janela principal (maior top-level do processo atual)
                string titleMain, classMain;
                var hMain = FindMainClientWindow(out titleMain, out classMain);
                if (hMain == IntPtr.Zero)
                    throw new InvalidOperationException("Não foi possível localizar a janela principal do cliente FTView neste processo.");

                Log($"Main HWND=0x{hMain.ToInt64():X}  Class='{classMain}'  Title='{titleMain}'");

                // 2) Acha o melhor filho (maior área cliente, com heurística)
                string titleChild, classChild;
                RECT rcClient;
                var hChild = FindLikelyDisplayChild(hMain, out titleChild, out classChild, out rcClient);

                var hCapture = hChild != IntPtr.Zero ? hChild : hMain; // fallback
                var used = (hChild != IntPtr.Zero) ? "CHILD" : "MAIN";

                // 3) Retângulo cliente → coordenadas de tela
                if (!GetClientRect(hCapture, out rcClient) || rcClient.Right <= rcClient.Left || rcClient.Bottom <= rcClient.Top)
                    throw new InvalidOperationException("GetClientRect falhou no HWND alvo.");

                POINT pt = new POINT { X = rcClient.Left, Y = rcClient.Top };
                if (!ClientToScreen(hCapture, ref pt))
                    throw new InvalidOperationException("ClientToScreen falhou no HWND alvo.");

                int w = rcClient.Right - rcClient.Left;
                int h = rcClient.Bottom - rcClient.Top;

                string titleUsed = (hChild != IntPtr.Zero) ? titleChild : titleMain;
                string classUsed = (hChild != IntPtr.Zero) ? classChild : classMain;

                Log($"Chosen [{used}] HWND=0x{hCapture.ToInt64():X}  Class='{classUsed}'  Title='{titleUsed}'  Client={w}x{h} @({pt.X},{pt.Y})");

                // 4) Screenshot da área cliente
                using (var bmp = new Bitmap(w, h))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(pt.X, pt.Y, 0, 0, new Size(w, h));
                    bmp.Save(outPngFullPath, ImageFormat.Png);
                }
                Log("Saved PNG: " + outPngFullPath);

                // 5) MsgBox (opcional)
                Mb("Processando (esperado): " + expectedPath +
                   "\nUsado: [" + used + "]" +
                   "\nHWND=0x" + hCapture.ToInt64().ToString("X") +
                   "\nClass='" + classUsed + "'" +
                   "\nTitle='" + titleUsed + "'" +
                   "\nClient=" + w + "x" + h +
                   "\nPNG: " + outPngFullPath);

                Log("=== ProcessDisplayByHwnd END ===");
                return "OK: " + outPngFullPath;
            }
            catch (Exception ex)
            {
                try
                {
                    var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}  [ERR] {ex.Message}";
                    File.AppendAllText(_diagLogPathHwnd, line + Environment.NewLine, Encoding.UTF8);
                }
                catch { }
                if (_diagShowMsgBoxHwnd)
                {
                    try { MessageBox.Show("ERRO: " + ex.Message + "\nVeja o log em:\n" + _diagLogPathHwnd, "FTViewCom/HWND", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                }
                return "#ERR: " + ex.Message;
            }
        }

        /// <summary>
        /// Dump textual de janelas top-level do processo e do maior filho da main (para debug rápido no VBA).
        /// </summary>
        [DispId(97)]
        public string DumpWindowsSnapshot()
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== DumpWindowsSnapshot ===");

            var list = EnumProcessTopLevelWindows(GetCurrentProcessId());
            sb.AppendLine($"Top-level count={list.Count}");
            foreach (var h in list)
            {
                var t = GetWindowTextSafe(h);
                var c = GetClassNameSafe(h);
                RECT r;
                GetWindowRect(h, out r);
                sb.AppendLine($"  HWND=0x{h.ToInt64():X}  Class='{c}'  Title='{t}'  Rect=({r.Left},{r.Top})-({r.Right},{r.Bottom})");
            }

            string tm, cm;
            var hMain = FindMainClientWindow(out tm, out cm);
            sb.AppendLine($"MainGuess HWND=0x{hMain.ToInt64():X}  Class='{cm}'  Title='{tm}'");

            string tch, cch; RECT rc;
            var hChild = (hMain != IntPtr.Zero) ? FindLikelyDisplayChild(hMain, out tch, out cch, out rc) : IntPtr.Zero;
/*           sb.AppendLine($"ChildGuess HWND=0x{hChild.ToInt64():X}  Class='{cch}'  Title='{tch}'  Client=({rc.Right - rc.Left}x{rc.Bottom - rc.Top})");
*/
            return sb.ToString();
        }

        // ===== Heurísticas de janela =====

        private static IntPtr FindMainClientWindow(out string title, out string klass)
        {
            title = ""; klass = "";
            var hwnds = EnumProcessTopLevelWindows(GetCurrentProcessId());
            IntPtr best = IntPtr.Zero; int bestArea = 0;

            foreach (var h in hwnds)
            {
                if (!IsWindowVisible(h)) continue;
                if (GetWindow(h, GW_OWNER) != IntPtr.Zero) continue;
                if (GetWindowRect(h, out RECT r))
                {
                    int area = Math.Max(0, r.Right - r.Left) * Math.Max(0, r.Bottom - r.Top);
                    if (area > bestArea)
                    {
                        bestArea = area;
                        best = h;
                    }
                }
            }

            if (best != IntPtr.Zero)
            {
                title = GetWindowTextSafe(best);
                klass = GetClassNameSafe(best);
            }
            return best;
        }

        private static IntPtr FindLikelyDisplayChild(IntPtr parent, out string title, out string klass, out RECT rcClient)
        {
            title = ""; klass = ""; rcClient = default;
            IntPtr best = IntPtr.Zero; int bestScore = 0;

            EnumChildWindows(parent, (h, l) =>
            {
                if (!IsWindowVisible(h)) return true;

                if (!GetClientRect(h, out RECT rc)) return true;
                int w = rc.Right - rc.Left, hh = rc.Bottom - rc.Top;
                if (w < 100 || hh < 100) return true;

                string cls = GetClassNameSafe(h);
                bool prefer = cls.IndexOf("FTV", StringComparison.OrdinalIgnoreCase) >= 0
                           || cls.IndexOf("HMI", StringComparison.OrdinalIgnoreCase) >= 0
                           || cls.IndexOf("Rna", StringComparison.OrdinalIgnoreCase) >= 0
                           || cls.IndexOf("View", StringComparison.OrdinalIgnoreCase) >= 0;

                int score = w * hh + (prefer ? 1_000_000 : 0);

                if (score > bestScore)
                {
                    bestScore = score;
                    best = h;
/*                    rcClient = rc;
                    title = GetWindowTextSafe(h);
                    klass = cls;*/
                }
                return true;
            }, IntPtr.Zero);

            return best;
        }

        // ===== Utils Win32 =====

        private static string GetWindowTextSafe(IntPtr hWnd)
        {
            var sb = new StringBuilder(512);
            try { GetWindowText(hWnd, sb, sb.Capacity); } catch { }
            return sb.ToString();
        }

        private static string GetClassNameSafe(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            try { GetClassName(hWnd, sb, sb.Capacity); } catch { }
            return sb.ToString();
        }

        private static System.Collections.Generic.List<IntPtr> EnumProcessTopLevelWindows(uint pid)
        {
            var list = new System.Collections.Generic.List<IntPtr>();
            EnumWindows((h, l) =>
            {
                GetWindowThreadProcessId(h, out uint wpid);
                if (wpid == pid && GetWindow(h, GW_OWNER) == IntPtr.Zero)
                    list.Add(h);
                return true;
            }, IntPtr.Zero);
            return list;
        }

        // P/Invoke
        [StructLayout(LayoutKind.Sequential)] private struct RECT { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)] private struct POINT { public int X, Y; }

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll", CharSet = CharSet.Auto)] private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
        private const uint GW_OWNER = 4;
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("kernel32.dll")] private static extern uint GetCurrentProcessId();
    }
}
