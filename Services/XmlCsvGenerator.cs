using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LibFTView.Services
{
    public sealed class XmlCsvGenerator
    {
        public sealed class Result
        {
            public string OutputFullPath { get; set; }
            public int RowCount { get; set; }
        }

        public Result Generate(string rootFolder, string outPath)
        {
            if (string.IsNullOrWhiteSpace(rootFolder))
                throw new ArgumentNullException(nameof(rootFolder));
            if (!Directory.Exists(rootFolder))
                throw new DirectoryNotFoundException(rootFolder);

            var rows = ListXmlRows(rootFolder);
            WriteCsv(outPath, rows);

            return new Result
            {
                OutputFullPath = Path.GetFullPath(outPath),
                RowCount = rows.Count
            };
        }

        private static List<(string Pasta, string Arquivo)> ListXmlRows(string root)
        {
            var files = Directory.EnumerateFiles(root, "*.xml", SearchOption.AllDirectories);

            var list = new List<(string Pasta, string Arquivo)>();
            foreach (var file in files)
            {
                var dirName = new DirectoryInfo(Path.GetDirectoryName(file) ?? string.Empty).Name;
                var fileNameNoExt = Path.GetFileNameWithoutExtension(file) ?? string.Empty;
                list.Add((dirName, fileNameNoExt));
            }

            return list
                .OrderBy(t => t.Pasta, StringComparer.OrdinalIgnoreCase)
                .ThenBy(t => t.Arquivo, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static void WriteCsv(string outPath, List<(string Pasta, string Arquivo)> rows)
        {
            var dir = Path.GetDirectoryName(outPath);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            var sb = new StringBuilder();
            sb.AppendLine("Pasta;Arquivo");

            foreach (var row in rows)
                sb.Append(Escape(row.Pasta)).Append(';').Append(Escape(row.Arquivo)).AppendLine();

            var utf8Bom = new UTF8Encoding(true);
            File.WriteAllText(outPath, sb.ToString(), utf8Bom);
        }

        private static string Escape(string value)
        {
            if (value == null) return "";
            bool needsQuotes = value.Contains(';') || value.Contains('"') || value.Contains('\n') || value.Contains('\r');
            if (needsQuotes)
            {
                var v = value.Replace("\"", "\"\"");
                return "\"" + v + "\"";
            }
            return value;
        }
    }
}
