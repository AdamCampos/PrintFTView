using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DisplayClient;

namespace LibFTView.Services
{
    [ComVisible(true)]
    [Guid("35C2E7C7-8F2C-4C67-9D7E-80B5E4388E6A")]
    [ProgId("LibFTView.FTVApp")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class FTVApp
    {
        // ===== LOG =====
        private static readonly object _logLock = new object();

        // 1) Mantém seu default...
        private string _logDir = @"C:\Projetos\VisualStudio\LibFTView";
        private string _logFile = "ftvapp.log";

        // ...mas tenta gravar em CommonAppData e TEMP se falhar.
        private void Log(string msg)
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}  {msg}";
            try
            {
                lock (_logLock)
                {
                    Directory.CreateDirectory(_logDir);
                    File.AppendAllText(Path.Combine(_logDir, _logFile), line + Environment.NewLine, new UTF8Encoding(false));
                }
            }
            catch (Exception ex1)
            {
                try
                {
                    var common = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "LibFTView");
                    Directory.CreateDirectory(common);
                    File.AppendAllText(Path.Combine(common, _logFile), "[FALLBACK CommonAppData] " + line + Environment.NewLine, new UTF8Encoding(false));
                }
                catch (Exception ex2)
                {
                    try
                    {
                        var tmp = Path.Combine(Path.GetTempPath(), "ftvapp.log");
                        File.AppendAllText(tmp, "[FALLBACK TEMP] " + line + $"  | e1={ex1.Message} | e2={ex2.Message}" + Environment.NewLine, new UTF8Encoding(false));
                    }
                    catch { /* ignora por completo */ }
                }
            }
        }

        [DispId(10)]
        public void SetLogPath(string folder, string fileName = "ftvapp.log")
        {
            if (!string.IsNullOrWhiteSpace(folder)) _logDir = folder;
            if (!string.IsNullOrWhiteSpace(fileName)) _logFile = fileName;
            try { Directory.CreateDirectory(_logDir); } catch { }
            Log($"[SetLogPath] dir='{_logDir}', file='{_logFile}'");
            try { MessageBox.Show($"Log em:\n{Path.Combine(_logDir, _logFile)}", "FTVApp.SetLogPath", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
        }

        // ===== DisplayClient.Application =====
        private global::DisplayClient.Application app;

        public FTVApp()
        {
            // construtor — se você não ver esta caixa, a ativação COM não aconteceu
            try { MessageBox.Show("FTVApp .ctor iniciado", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }

            try { Directory.CreateDirectory(_logDir); } catch { }
            var asm = Assembly.GetExecutingAssembly();
            Log("=== FTVApp .ctor ===");
            Log("[Assembly] " + asm.FullName);
            Log("[Bitness] " + (IntPtr.Size == 8 ? "x64" : "x86"));
            Log("[Process] PID=" + System.Diagnostics.Process.GetCurrentProcess().Id);

            try { MessageBox.Show("FTVApp .ctor OK", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
        }

        private bool EnsureApp()
        {
            if (app != null)
            {
                Log("[EnsureApp] Reutilizando DisplayClient.Application");
                try { MessageBox.Show("EnsureApp: reutilizando instancia", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return true;
            }

            Log("[EnsureApp] Criando DisplayClient.Application...");
            try
            {
                try { MessageBox.Show("EnsureApp: criando DisplayClient.Application...", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                app = new global::DisplayClient.Application();
                Log("[EnsureApp] OK");
                try { MessageBox.Show("EnsureApp: OK", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return true;
            }
            catch (COMException comEx)
            {
                Log($"[EnsureApp][COMEX] 0x{comEx.ErrorCode:X} {comEx.Message}");
                try { MessageBox.Show($"Erro ao inicializar DisplayClient.Application:\n{comEx.Message}\n(0x{comEx.ErrorCode:X})", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                return false;
            }
            catch (Exception ex)
            {
                Log("[EnsureApp][EX] " + ex);
                try { MessageBox.Show($"Erro inesperado em EnsureApp:\n{ex}", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                return false;
            }
        }

        [DispId(1)]
        public string OpenTopsideWellA()
            => OpenDisplay("/TOPSIDE::l3_1210_10001_sat_prod_well_a", "");

        [DispId(2)]
        public string OpenDisplay(string nomeTela, string parametro)
        {
            Log($"[OpenDisplay] in display='{nomeTela}' param='{parametro}'");

            var display = (nomeTela ?? string.Empty).Trim();
            var param = parametro ?? string.Empty;
            if (string.IsNullOrWhiteSpace(display)) return "#ERR: nomeTela vazio";
            if (!EnsureApp()) return "#ERR: DisplayClient.Application não inicializado";

            try
            {
                Log("[OpenDisplay] LoadDisplay/ShowDisplay...");
                app.LoadDisplay(display, param);
                app.ShowDisplay(display, param);

                Log("[OpenDisplay] sucesso -> " + display);
                try { MessageBox.Show("OpenDisplay OK:\n" + display, "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
                return "OK: " + display + (param.Length > 0 ? $" ({param})" : "");
            }
            catch (COMException comEx)
            {
                Log($"[OpenDisplay][COMEX] 0x{comEx.ErrorCode:X} {comEx.Message}");
                try { MessageBox.Show($"OpenDisplay COMEX:\n0x{comEx.ErrorCode:X}\n{comEx.Message}", "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                return $"#ERR COM 0x{comEx.ErrorCode:X}: {comEx.Message}";
            }
            catch (Exception ex)
            {
                Log("[OpenDisplay][EX] " + ex);
                try { MessageBox.Show("OpenDisplay EX:\n" + ex, "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Error); } catch { }
                return "#ERR: " + ex.Message;
            }
        }

        [DispId(3)]
        public string Ping()
        {
            Log("[Ping] start");
            var ok = EnsureApp();
            Log("[Ping] end => " + (ok ? "OK" : "FAIL"));
            try { MessageBox.Show("Ping: " + (ok ? "OK" : "FAIL"), "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
            return ok ? "OK" : "#ERR";
        }
    }
}
