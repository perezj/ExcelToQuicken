using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BanortePDFToExcel
{
    internal class Util
    {

        internal static bool IsHeaderLine(string line)
        {
            var normalized = Normalize(line);

            return normalized.Contains("FECHA") &&
                   normalized.Contains("DESCRIPCIÓN") &&
                   normalized.Contains("MONTO") &&
                   normalized.Contains("SALDO");
        }
        internal static string Normalize(string text)
        {
            return Regex.Replace(text.ToUpper(), @"\s+", " ").Trim();
        }
    }
}
