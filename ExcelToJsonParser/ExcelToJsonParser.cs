using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using ExcelDataReader;
using NJsonSchema.CodeGeneration.CSharp;
using Spire.Xls;

namespace Rochas.ExcelToJson
{
    public static class ExcelToJsonParser
    {
        #region Public Methods

        public static string GetJsonString(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var counter = 0;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var fileStream = GetFileStream(fileName))
            {
                using (var result = new StringWriter())
                {
                    var readerConfig = new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252)
                    };
                    using (var reader = ExcelReaderFactory.CreateReader(fileStream, readerConfig))
                    {
                        using (var writer = new JsonTextWriter(result))
                        {
                            writer.Formatting = Formatting.Indented;
                            writer.WriteStartArray();

                            while (skipRows > 0)
                            {
                                reader.Read();
                                skipRows--;
                            }

                            reader.Read();

                            if (headerColumns == null)
                                headerColumns = GetHeaderColumns(reader);
                            else
                            {
                                if (headerColumns.Length < reader.FieldCount)
                                    throw new Exception("Invalid column amount");
                            }

                            ApplyColumnNamesReplace(headerColumns, replaceFrom, replaceTo);

                            do
                            {
                                while (reader.Read() && (!onlySampleRow || (onlySampleRow && counter < 1)))
                                {
                                    WriteItemJsonBody(reader, writer, headerColumns);
                                    counter += 1;
                                }

                            } while (reader.NextResult());

                            writer.WriteEndArray();
                        }
                    }

                    return result.ToString();
                }
            }
        }

        public static IEnumerable<object> GetJsonObject(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var strJson = GetJsonString(fileName, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);

            return JsonConvert.DeserializeObject<IEnumerable<object>>(strJson);
        }

        public static string GetJsonClassModel(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null)
        {
            string result = null;
            var jsonContent = GetJsonString(fileName, skipRows, replaceFrom, replaceTo, headerColumns, true);
            if (!string.IsNullOrWhiteSpace(jsonContent))
            {
                var schema = NJsonSchema.JsonSchema.FromSampleJson(jsonContent);

                var genOptions = new CSharpGeneratorSettings()
                {
                    GenerateDataAnnotations = false,
                    GenerateDefaultValues = false,
                    GenerateJsonMethods = true
                };
                var generator = new CSharpGenerator(schema, genOptions);

                var className = fileName.Replace(".xlsx", string.Empty).Replace(".xls", string.Empty);

                result = generator.GenerateFile(className);
            }

            return result;
        }

        public static DataTable GetDataTable(string fileName, int skipRows = 0, bool useHeader = true)
        {
            using (var fileStream = GetFileStream(fileName))
            {
                var reader = ExcelReaderFactory.CreateReader(fileStream);

                var config = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = useHeader
                    }
                };
                
                while (skipRows > 0)
                {
                    reader.Read();
                    skipRows--;
                }

                return reader.AsDataSet(config).Tables[0];
            }
        }

        #endregion

        #region Helper Methods

        private static Stream GetFileStream(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("File name not informed");

            var isBinaryFile = Path.GetExtension(fileName).ToLower().Equals(".xlsb");
            if (!isBinaryFile)
                return File.Open(fileName, FileMode.Open, FileAccess.Read);
            else
            {
                var result = new MemoryStream();

                Workbook workbook = new Workbook();
                workbook.LoadFromFile(fileName);
                workbook.SaveToStream(result);

                return result;
            }
        }

        private static string[] GetHeaderColumns(IExcelDataReader reader)
        {
            var result = new string[reader.FieldCount];

            for (var count = 0; count < reader.FieldCount; count++)
                result[count] = reader[count].ToString();

            return result;
        }

        private static void ApplyColumnNamesReplace(string[] columnNames, string[] readFrom, string[] replaceTo)
        {
            if ((readFrom != null) && (replaceTo != null))
            {
                if (readFrom.Length != replaceTo.Length)
                    throw new ArgumentOutOfRangeException("Invalid replace values amount");

                for (var nameCount = 0; nameCount < columnNames.Length; nameCount++)
                {
                    for (var chrCount = 0; chrCount < readFrom.Length; chrCount++)
                        columnNames[nameCount] = columnNames[nameCount].Replace(readFrom[chrCount], replaceTo[chrCount]).Replace(" ", "_");
                }
            }
        }

        private static void WriteItemJsonBody(IExcelDataReader reader, JsonWriter writer, string[] headerColumns)
        {
            writer.WriteStartObject();

            var colCount = 0;
            foreach (var col in headerColumns)
            {
                var colValue = reader.GetValue(colCount);
                writer.WritePropertyName(col);
                writer.WriteValue(colValue);
                colCount += 1;
            }

            writer.WriteEndObject();
        }

        #endregion
    }
}
