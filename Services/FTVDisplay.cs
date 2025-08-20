using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DisplayClient; // referência COM
using System.Reflection;

namespace LibFTView.Services
{
    [System.Runtime.InteropServices.ComVisible(true)]
    [System.Runtime.InteropServices.Guid("B7A47D2B-752B-4A8A-8B3A-07E7C93E5A5C")]
    [System.Runtime.InteropServices.ProgId("LibFTView.FTVDisplay")]
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.AutoDual)]
    public class FTVDisplay
    {
        private DisplayClient.Application app;

        public FTVDisplay()
        {
            // garanta que o log vá para a pasta exigida
            DiagLog.SetPath(@"C:\Projetos\VisualStudio\LibFTView");
            DiagLog.Write("FTVDisplay .ctor — assembly=" + Assembly.GetExecutingAssembly().FullName);
        }

        [System.Runtime.InteropServices.DispId(1)]
        public string OpenTopsideDetectorsVoting()
            => OpenDisplay(@"Topside\detectors_voting_group_fire_1_4", null);

        [System.Runtime.InteropServices.DispId(2)]
        public string OpenDisplay(string nomeTela, string parametro)
        {
            DiagLog.Write($"OpenDisplay(nomeTela='{nomeTela}', parametro='{parametro}') — start");

            if (string.IsNullOrWhiteSpace(nomeTela))
            {
                DiagLog.Write("ERRO: nomeTela vazio");
                return "#ERR: nomeTela vazio";
            }

            try
            {
                if (app is null)
                {
                    try
                    {
                        DiagLog.Write("Inicializando DisplayClient.Application...");
                        app = new DisplayClient.Application();
                        DiagLog.Write("DisplayClient.Application OK");
                    }
                    catch (COMException comEx)
                    {
                        DiagLog.Write($"COMEX ao instanciar DisplayClient.Application: 0x{comEx.ErrorCode:X} {comEx.Message}");
                        MessageBox.Show($"Erro ao inicializar DisplayClient.Application: {comEx.Message} (Código: {comEx.ErrorCode:X})");
                        return $"#ERR COM 0x{comEx.ErrorCode:X}: {comEx.Message}";
                    }
                    catch (Exception ex)
                    {
                        DiagLog.Write("EX ao instanciar DisplayClient.Application: " + ex);
                        MessageBox.Show($"Erro inesperado ao inicializar DisplayClient.Application: {ex.Message}");
                        return "#ERR: " + ex.Message;
                    }
                }

                try
                {
                    if (!string.IsNullOrEmpty(parametro))
                    {
                        DiagLog.Write("LoadDisplay + ShowDisplay (com parâmetro)...");
                        app.LoadDisplay(nomeTela, parametro);
                        app.ShowDisplay(nomeTela, parametro);
                    }
                    else
                    {
                        DiagLog.Write("LoadDisplay + ShowDisplay (sem parâmetro)...");
                        app.LoadDisplay(nomeTela);
                        app.ShowDisplay(nomeTela);
                    }
                    DiagLog.Write("OpenDisplay — sucesso");
                    return "OK: " + nomeTela + (string.IsNullOrEmpty(parametro) ? "" : $" ({parametro})");
                }
                catch (COMException comEx)
                {
                    DiagLog.Write($"COMEX em Show/Load: 0x{comEx.ErrorCode:X} {comEx.Message}");
                    return $"#ERR COM 0x{comEx.ErrorCode:X}: {comEx.Message}";
                }
                catch (Exception ex)
                {
                    DiagLog.Write("EX em Show/Load: " + ex);
                    return "#ERR: " + ex.Message;
                }
            }
            finally
            {
                DiagLog.Write("OpenDisplay — end");
            }
        }

        // (Opcional) Se você usar também FTViewCom.ProcessDisplayByHwnd:
        [System.Runtime.InteropServices.DispId(90)]
        public void RedirectFTViewComLog()
        {
            try
            {
                // redireciona o log interno do FTViewCom para a mesma pasta
                new FTViewCom().SetHwndDiagOptions(
                    showMessageBoxes: true,
                    logPath: @"C:\Projetos\VisualStudio\LibFTView\ftviewcom_hwnd.log"
                );
                DiagLog.Write("FTViewCom log redirecionado para ftviewcom_hwnd.log");
            }
            catch (Exception ex)
            {
                DiagLog.Write("Falha ao redirecionar FTViewCom log: " + ex.Message);
            }
        }
    }
}
