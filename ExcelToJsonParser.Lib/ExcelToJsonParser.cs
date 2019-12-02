using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using ExcelDataReader;
using NJsonSchema.CodeGeneration.CSharp;

namespace Rochas.ExcelToJson.Lib
{
    public static class ExcelToJsonParser
    {
        #region Public Methods

        public static string GetJsonString(string fileName, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var counter = 0;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var inputFile = File.Open(fileName, FileMode.Open, FileAccess.Read))
            using (var result = new StringWriter())
            {
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };
                using (var reader = ExcelReaderFactory.CreateReader(inputFile, readerConfig))
                {
                    using (var writer = new JsonTextWriter(result))
                    {
                        writer.Formatting = Formatting.Indented;
                        writer.WriteStartArray();

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
                            while (reader.Read() && (onlySampleRow && counter < 1))
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

        public static IEnumerable<object> GetJsonObject(string fileName, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var strJson = GetJsonString(fileName, replaceFrom, replaceTo, headerColumns, onlySampleRow);

            return JsonConvert.DeserializeObject<IEnumerable<object>>(strJson);
        }

        public static string GetJsonClassModel(string fileName, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null)
        {
            string result = null;
            var jsonContent = GetJsonString(fileName, replaceFrom, replaceTo, headerColumns, true);
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

        public static DataSet GetDataTable(string fileName)
        {
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelReaderFactory.CreateReader(stream);

                var config = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };
                return reader.AsDataSet(config);
            }

        }

        #endregion

        #region Helper Methods

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
