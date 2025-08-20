using System;
using LibFTView.Services;

namespace LibFTView
{
    /// <summary>
    /// CLI entry point.
    /// Usage:
    ///   LibFTView.exe ftv [--log "C:\Temp\ftvapp.log"] [--open "/TOPSIDE::..."]
    ///   LibFTView.exe listaxml [--dir "C:\RSLogix 5000\Projects\P83\Telas\TELAS_XML\GFX"]
    /// </summary>
    internal static class Program
    {
        private static int Main(string[] args)
        {
            if (args == null || args.Length == 0 || Array.Exists(args, a => a == "--help" || a == "-h" || a == "/?"))
            {
                ShowHelp();
                return 1;
            }

            var cmd = args[0].ToLowerInvariant();
            try
            {
                switch (cmd)
                {
                    case "ftv":
                        return RunFtv(args);
                    case "listaxml":
                        return RunListaXml(args);
                    default:
                        Console.Error.WriteLine("Comando inválido: " + cmd);
                        ShowHelp();
                        return 2;
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[FATAL] " + ex);
                return 99;
            }
        }

        private static int RunFtv(string[] args)
        {
            string logPath = null;
            string display = null;
            string parametro = "";

            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "--log" && i + 1 < args.Length) logPath = args[++i];
                else if (args[i] == "--open" && i + 1 < args.Length) display = args[++i];
                else if (args[i] == "--param" && i + 1 < args.Length) parametro = args[++i];
            }

            var core = new FTVCore();
            if (!string.IsNullOrWhiteSpace(logPath))
            {
                var folder = System.IO.Path.GetDirectoryName(logPath);
                var file = System.IO.Path.GetFileName(logPath);
                core.SetLogPath(folder, string.IsNullOrWhiteSpace(file) ? "ftvapp.log" : file);
            }

            Console.WriteLine(core.Ping());

            if (string.IsNullOrWhiteSpace(display))
                Console.WriteLine(core.OpenTopsideWellA());
            else
                Console.WriteLine(core.OpenDisplay(display, parametro));

            return 0;
        }

        private static int RunListaXml(string[] args)
        {
            string dir = null;
            for (int i = 1; i < args.Length; i++)
            {
                if (args[i] == "--dir" && i + 1 < args.Length) dir = args[++i];
            }

            var leitor = new ListaXML();
            var itens = leitor.Enumerar(dir);

            // imprime resultado em stdout
            Console.WriteLine("AREA;TELA");
            foreach (var item in itens)
            {
                Console.WriteLine($"{item.Area};{item.Tela}");
            }

            Console.Error.WriteLine($"[ListaXML] Diretório raiz: {leitor.UltimoDiretorioRaiz}");
            Console.Error.WriteLine($"[ListaXML] Total de XMLs: {itens.Count}");

            return 0;
        }

        private static void ShowHelp()
        {
            Console.WriteLine(@"Uso:
  LibFTView.exe ftv [--log ""C:\Temp\ftvapp.log""] [--open ""/TOPSIDE::l3_1210_10001_sat_prod_well_a""] [--param """"]
  LibFTView.exe listaxml [--dir ""C:\RSLogix 5000\Projects\P83\Telas\TELAS_XML\GFX""]

Observações:
- 'ftv' usa a classe FTVCore (sem COM) e mantém compatibilidade via FTVApp (COM) para VBA.
- 'listaxml' lista arquivos XML em GFX\HULL e GFX\TOPSIDE gerando pares AREA;TELA.");
        }
    }
}
