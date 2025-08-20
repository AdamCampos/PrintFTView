using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace LibFTView.Services
{
    /// <summary>
    /// Varre GFX\HULL e GFX\TOPSIDE e devolve pares (AREA, TELA).
    /// AREA = HULL/TOPSIDE; TELA = nome do arquivo XML sem extensão.
    /// </summary>
    public class ListaXML
    {
        public string UltimoDiretorioRaiz { get; private set; }

        public sealed class Item
        {
            public string Area { get; set; }
            public string Tela { get; set; }
        }

        /// <summary>
        /// Enumera XMLs no diretório informado (ou default) e retorna lista (AREA, TELA).
        /// </summary>
        /// <param name="raiz">Diretório base GFX. Se vazio/nulo, usa o default.</param>
        public List<Item> Enumerar(string raiz = null)
        {
            var defaultDir = @"C:\RSLogix 5000\Projects\P83\Telas\TELAS_XML\GFX";
            raiz = string.IsNullOrWhiteSpace(raiz) ? defaultDir : raiz.Trim();

            // Se apontar direto pra HULL/TOPSIDE, sobe um nível (GFX)
            var di = new DirectoryInfo(raiz);
            if (di.Exists && (di.Name.Equals("HULL", StringComparison.OrdinalIgnoreCase) ||
                              di.Name.Equals("TOPSIDE", StringComparison.OrdinalIgnoreCase)))
            {
                raiz = di.Parent?.FullName ?? raiz;
            }

            if (!Directory.Exists(raiz))
                throw new DirectoryNotFoundException($"Diretório não encontrado: {raiz}");

            UltimoDiretorioRaiz = raiz;

            var resultado = new List<Item>();
            foreach (var area in new[] { "HULL", "TOPSIDE" })
            {
                var sub = Path.Combine(raiz, area);
                if (!Directory.Exists(sub)) continue;

                foreach (var arq in Directory.EnumerateFiles(sub, "*.xml", SearchOption.TopDirectoryOnly))
                {
                    var nome = Path.GetFileNameWithoutExtension(arq);
                    resultado.Add(new Item { Area = area, Tela = nome });
                }
            }

            return resultado
                .OrderBy(i => i.Area, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.Tela, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
    }
}
