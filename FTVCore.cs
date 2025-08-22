using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using DisplayClient;
using LibFTView.Win32;
using LibFTView.Services;

namespace LibFTView.Services
{
    /// <summary>
    /// Núcleo sem COM/MessageBox. Pode ser chamado via CLI ou por camadas superiores.
    /// </summary>
    public class FTVCore
    {
        private static readonly object _logLock = new object();
        private string _logDir = @"C:\Projetos\VisualStudio\LibFTView";
        private string _logFile = "ftvapp.log";
        private DisplayClient.Application app;

        public FTVCore()
        {
            try { Directory.CreateDirectory(_logDir); } catch { }
            var asm = Assembly.GetExecutingAssembly();
            Log("=== FTVCore .ctor ===");
            Log("[Assembly] " + asm.FullName);
            Log("[Bitness] " + (IntPtr.Size == 8 ? "x64" : "x86"));
            Log("[Process] PID=" + System.Diagnostics.Process.GetCurrentProcess().Id);
        }

        public void SetLogPath(string folder, string fileName = "ftvapp.log")
        {
            if (!string.IsNullOrWhiteSpace(folder)) _logDir = folder;
            if (!string.IsNullOrWhiteSpace(fileName)) _logFile = fileName;
            try { Directory.CreateDirectory(_logDir); } catch { }
            Log($"[SetLogPath] dir='{_logDir}', file='{_logFile}'");
        }

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
                    catch { }
                }
            }
        }

        private bool EnsureApp()
        {
            if (app != null)
            {
                Log("[EnsureApp] Reutilizando DisplayClient.Application");
                return true;
            }

            Log("[EnsureApp] Criando DisplayClient.Application...");
            try
            {
                app = new DisplayClient.Application();
                Log("[EnsureApp] OK");
                return true;
            }
            catch (COMException comEx)
            {
                Log($"[EnsureApp][COMEX] 0x{comEx.ErrorCode:X} {comEx.Message}");
                return false;
            }
            catch (Exception ex)
            {
                Log("[EnsureApp][EX] " + ex);
                return false;
            }
        }

        public string OpenTopsideWellA()
            => OpenDisplay("/TOPSIDE::l3_1210_10001_sat_prod_well_a", "");

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

                // === Captura (print) da janela/tela aberta ===
                try
                {
                    System.Threading.Thread.Sleep(200); // dar tempo para o child aparecer

                    if (WindowProbe.TryGetFtViewRenderTarget(out var capHwnd, out var fullFromTitle, Log))
                    {
                        // se conseguimos "/AREA::NOME" a partir do título, priorizar
                        var fullForFile = !string.IsNullOrWhiteSpace(fullFromTitle) ? fullFromTitle : display;

                        Log($"[PRINT] Preparando captura display='{fullForFile}'");
                        var savedPath = DisplayPrinter.CaptureAfterStable(
                                            capHwnd,
                                            fullForFile,
                                            DisplayPrinter.DefaultOutDir,
                                            maxWaitMs: 7000,
                                            sampleIntervalMs: 200,
                                            requiredStableSamples: 3,
                                            log: Log);

                        if (!string.IsNullOrEmpty(savedPath))
                            Log("[PRINT] OK file='" + savedPath + "'");
                        else
                            Log("[PRINT] NÃO FOI SALVO (timeout/falha).");
                    }
                    else
                    {
                        Log("[PRINT] Nenhum HWND renderizador encontrado; snapshot de diagnóstico:");
                        WindowProbe.SnapshotFtView(display, Log, childLimit: 8);
                    }
                }
                catch (Exception ex)
                {
                    Log("[PRINT][EX] " + ex);
                }
                // === fim captura ===

                // Log detalhado de HWND/container/filhos e tipagem (Janela/Tela)
                try { WindowProbe.SnapshotFtView(display, Log, childLimit: 8); }
                catch (Exception ex) { Log("[HWND][FTView][EX] " + ex.Message); }

                // fecha automaticamente após 3s (se desejado)
                System.Threading.Tasks.Task.Delay(3000).ContinueWith(_ =>
                {
                    try
                    {
                        Log("[OpenDisplay] UnloadDisplay -> " + display);
                        app.UnloadDisplay(display);
                    }
                    catch (Exception ex)
                    {
                        Log("[OpenDisplay][UnloadDisplay][EX] " + ex);
                    }
                });

                Log("[OpenDisplay] sucesso -> " + display);
                return "OK: " + display + (param.Length > 0 ? $" ({param})" : "");
            }
            catch (COMException comEx)
            {
                Log($"[OpenDisplay][COMEX] 0x{comEx.ErrorCode:X} {comEx.Message}");
                return $"#ERR COM 0x{comEx.ErrorCode:X}: {comEx.Message}";
            }
            catch (Exception ex)
            {
                Log("[OpenDisplay][EX] " + ex);
                return "#ERR: " + ex.Message;
            }
        }

        public string Ping()
        {
            Log("[Ping] start");
            var ok = EnsureApp();
            Log("[Ping] end => " + (ok ? "OK" : "FAIL"));
            return ok ? "OK" : "#ERR";
        }
    }
}
