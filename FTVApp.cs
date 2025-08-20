using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using DisplayClient;
using System.Linq;


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

        // 1) Default do log
        private string _logDir = @"C:\Projetos\VisualStudio\LibFTView";
        private string _logFile = "ftvapp.log";

        // Fallback de log (CommonAppData e TEMP) caso falhe no diretório padrão
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
        private DisplayClient.Application app;

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
                app = new DisplayClient.Application();
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

                // fecha automaticamente após 3 segundos
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


        [DispId(3)]
        public string Ping()
        {
            Log("[Ping] start");
            var ok = EnsureApp();
            Log("[Ping] end => " + (ok ? "OK" : "FAIL"));
            try { MessageBox.Show("Ping: " + (ok ? "OK" : "FAIL"), "FTVApp", MessageBoxButtons.OK, MessageBoxIcon.Information); } catch { }
            return ok ? "OK" : "#ERR";
        }

        // ====================== NOVAS FUNÇÕES: LISTAGEM DE XMLs =========================

        [DispId(20)]
        public string ListaXML_CSV(string raiz)
        {
            try
            {
                var leitor = new ListaXML();
                List<ListaXML.Item> itens = leitor.Enumerar(string.IsNullOrWhiteSpace(raiz) ? null : raiz);

                // monta CSV (AREA;TELA)
                var sb = new StringBuilder();
                sb.AppendLine("AREA;TELA");
                foreach (var it in itens)
                    sb.AppendLine($"{it.Area};{it.Tela}");

                // loga diretório e contagem
                Log($"[ListaXML] Diretório raiz: {leitor.UltimoDiretorioRaiz}");
                Log($"[ListaXML] Total de XMLs: {itens.Count}");

                return sb.ToString();
            }
            catch (Exception ex)
            {
                Log("[ListaXML_CSV][EX] " + ex);
                return "#ERR: " + ex.Message;
            }
        }

        [DispId(21)]
        public object ListaXML_Array(string raiz)
        {
            try
            {
                var leitor = new ListaXML();
                List<ListaXML.Item> itens = leitor.Enumerar(string.IsNullOrWhiteSpace(raiz) ? null : raiz);

                // cria SAFEARRAY 1-based (Excel-friendly): [1..rows, 1..cols]
                int rows = itens.Count + 1; // +1 cabeçalho
                int cols = 2;
                Array arr = Array.CreateInstance(typeof(object), new int[] { rows, cols }, new int[] { 1, 1 });

                // header
                arr.SetValue("AREA", 1, 1);
                arr.SetValue("TELA", 1, 2);

                // dados
                int r = 2;
                foreach (var it in itens)
                {
                    arr.SetValue(it.Area, r, 1);
                    arr.SetValue(it.Tela, r, 2);
                    r++;
                }

                // loga diretório e contagem
                Log($"[ListaXML] Diretório raiz: {leitor.UltimoDiretorioRaiz}");
                Log($"[ListaXML] Total de XMLs: {itens.Count}");

                return arr; // VBA pode jogar direto em Range.Value2
            }
            catch (Exception ex)
            {
                Log("[ListaXML_Array][EX] " + ex);
                return null;
            }
        }

        [DispId(22)]
        public string ListaXML_SaveCSV(string raiz, string caminhoCSV)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(caminhoCSV))
                    return "#ERR: caminhoCSV vazio";

                var leitor = new ListaXML();
                var itens = leitor.Enumerar(string.IsNullOrWhiteSpace(raiz) ? null : raiz);

                var dir = Path.GetDirectoryName(caminhoCSV);
                if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

                using (var sw = new StreamWriter(caminhoCSV, false, new UTF8Encoding(false)))
                {
                    sw.WriteLine("AREA;TELA");
                    foreach (var it in itens) sw.WriteLine($"{it.Area};{it.Tela}");
                }

                Log($"[ListaXML] Diretório raiz: {leitor.UltimoDiretorioRaiz}");
                Log($"[ListaXML] Total de XMLs: {itens.Count}");
                Log($"[ListaXML] CSV salvo em: {caminhoCSV}");

                return $"OK; count={itens.Count}; dir={leitor.UltimoDiretorioRaiz}";
            }
            catch (Exception ex)
            {
                Log("[ListaXML_SaveCSV][EX] " + ex);
                return "#ERR: " + ex.Message;
            }
        }


        // === Helpers internos para ler CSV "AREA;TELA" ===
        private IEnumerable<(string Area, string Tela)> ParseListaCsv(string path)
        {
            foreach (var raw in File.ReadAllLines(path, new UTF8Encoding(false)))
            {
                var line = (raw ?? "").Trim();
                if (line.Length == 0) continue;
                if (line.StartsWith("#")) continue;
                if (line.StartsWith("AREA;", StringComparison.OrdinalIgnoreCase)) continue;

                var parts = line.Split(';');
                if (parts.Length < 2) continue;

                var area = (parts[0] ?? "").Trim();
                var tela = (parts[1] ?? "").Trim();
                if (area.Length == 0 || tela.Length == 0) continue;

                yield return (area, tela);
            }
        }

        // === Núcleo: abre todas as telas listadas, aguardando entre elas ===
        private string AbrirTelasCore(IEnumerable<(string Area, string Tela)> itens, int waitMsPorTela, string areaFiltro, string parametro)
        {
            var lista = itens?.ToList() ?? new List<(string Area, string Tela)>();
            if (!string.IsNullOrWhiteSpace(areaFiltro))
                lista = lista.Where(i => i.Area.Equals(areaFiltro, StringComparison.OrdinalIgnoreCase)).ToList();

            if (!EnsureApp())
                return "#ERR: DisplayClient.Application não inicializado";

            int total = 0, ok = 0, err = 0;

            foreach (var it in lista)
            {
                var display = $"/{it.Area}::{it.Tela}";
                try
                {
                    Log($"[AbrirTelas] {display} param='{parametro}'");
                    app.LoadDisplay(display, parametro ?? "");
                    app.ShowDisplay(display, parametro ?? "");
                    total++;
                    ok++;
                    if (waitMsPorTela > 0)
                        System.Threading.Thread.Sleep(waitMsPorTela);
                }
                catch (COMException comEx)
                {
                    err++;
                    Log($"[AbrirTelas][COMEX] display='{display}' 0x{comEx.ErrorCode:X} {comEx.Message}");
                }
                catch (Exception ex)
                {
                    err++;
                    Log($"[AbrirTelas][EX] display='{display}' {ex}");
                }
            }

            var resumo = $"OK; total={total}; ok={ok}; err={err}";
            Log("[AbrirTelas] " + resumo);
            return resumo;
        }

        // === API COM única (aceita diretório GFX OU arquivo CSV) ===
        [DispId(30)]
        public string AbrirTelas(string origem, int waitMsPorTela = 1000, string area = "", string parametro = "")
        {
            try
            {
                Log($"[AbrirTelas] origem='{origem}' waitMs={waitMsPorTela} area='{area}' param='{parametro}'");

                // Caso seja CSV existente -> lê CSV; caso contrário -> enumera diretório (GFX)
                if (!string.IsNullOrWhiteSpace(origem) &&
                    File.Exists(origem) &&
                    string.Equals(Path.GetExtension(origem), ".csv", StringComparison.OrdinalIgnoreCase))
                {
                    var itensCsv = ParseListaCsv(origem);
                    return AbrirTelasCore(itensCsv, waitMsPorTela, area, parametro);
                }
                else
                {
                    var leitor = new ListaXML();
                    var itens = leitor.Enumerar(string.IsNullOrWhiteSpace(origem) ? null : origem)
                                      .Select(i => (i.Area, i.Tela));
                    Log($"[AbrirTelas] Diretório raiz: {leitor.UltimoDiretorioRaiz}");
                    return AbrirTelasCore(itens, waitMsPorTela, area, parametro);
                }
            }
            catch (Exception ex)
            {
                Log("[AbrirTelas][EX] " + ex);
                return "#ERR: " + ex.Message;
            }
        }

        // === Conveniências COM explícitas (se preferir separar CSV de DIR no VBA) ===
        [DispId(31)]
        public string AbrirTelasDeDiretorio(string raizGfx, int waitMsPorTela = 1000, string area = "", string parametro = "")
            => AbrirTelas(raizGfx, waitMsPorTela, area, parametro);

        [DispId(32)]
        public string AbrirTelasDeCSV(string caminhoCSV, int waitMsPorTela = 1000, string area = "", string parametro = "")
            => AbrirTelas(caminhoCSV, waitMsPorTela, area, parametro);


    }


}
