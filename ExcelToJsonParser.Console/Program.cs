using System;

namespace Rochas.ExcelToJson
{
    class Program
    {
        static void Main(string[] args)
        {
            var replaceFrom = new string[20] { "á", "à", "ã", "é", "ê", "í", "ó", "ú", "Á", "À", "Ã", "É", "Ê", "Í", "Ó", "Ú", "ç", "Ç", "R$", "%" };
            var replaceTo = new string[20] { "a", "a", "a", "e", "e", "i", "o", "u", "A", "A", "A", "E", "E", "I", "O", "U", "c", "C", "Reais", "Perc" };

            var excelContent = ExcelToJsonParser.GetJsonString("TesteFinanc.xlsx", replaceFrom:replaceFrom, replaceTo:replaceTo);
            Console.WriteLine(excelContent);
            Console.Read();
        }
    }
}
