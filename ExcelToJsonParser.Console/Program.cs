using System;
using System.Threading;

namespace Rochas.ExcelToJson
{
    class Program
    {
        static readonly string[] _replaceFrom = new string[20] { "á", "à", "ã", "é", "ê", "í", "ó", "ú", "Á", "À", "Ã", "É", "Ê", "Í", "Ó", "Ú", "ç", "Ç", "R$", "%" };
        static readonly string[] _replaceTo = new string[20] { "a", "a", "a", "e", "e", "i", "o", "u", "A", "A", "A", "E", "E", "I", "O", "U", "c", "C", "Reais", "Perc" };

        static void Main(string[] args)
        {
            try
            {
                var skipLines = 0;
                var excelContent = ExcelToJsonParser.GetJsonStringFromTabular("Samples\\TabularSample.xlsx", skipLines, _replaceFrom, _replaceTo);
                Console.Clear();
                Console.WriteLine("Tabular Sheet Result Sample :");
                Console.WriteLine(excelContent);
                Console.Read();

                excelContent = ExcelToJsonParser.GetJsonStringFromForm("Samples\\FormSample.xlsx", "PlanTeste1", _replaceFrom, _replaceTo);
                Console.Clear();
                Console.WriteLine("Form Sheet Result Sample :");
                Console.WriteLine(excelContent);
                Console.Read();
                Console.Read();
            }
            catch(Exception ex)
            {
                Console.WriteLine($"An error ocurred while parsing excel file content:{ex.Message}{Environment.NewLine}{ex.StackTrace}");
                Console.Read();
            }
        }
    }
}
