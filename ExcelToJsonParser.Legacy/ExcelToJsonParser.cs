using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using ExcelDataReader; // Used to auto parse sheet data to datareader
using ClosedXML.Excel; // Used to access named cells
using Spire.Xls; // Used only to open excel binary files
using NJsonSchema.CodeGeneration.CSharp; // Used to generate C# class model code

namespace Rochas.ExcelToJson
{
    public static class ExcelToJsonParser
    {
        #region Tabular Sheet Parser Public Methods

        public static string GetJsonStringFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            using (var fileContent = GetFileStream(fileName))
            {
                return GetJsonStringFromTabular(fileContent, skipRows, replaceFrom, replaceTo, headerColumns);
            }
        }

        public static string GetJsonStringFromTabular(Stream fileContent, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var counter = 0;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);


            using (var result = new StringWriter())
            {
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };
                using (var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig))
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
                                WriteItemJsonBodyFromReader(reader, writer, headerColumns);
                                counter += 1;
                            }

                        } while (reader.NextResult());

                        writer.WriteEndArray();
                    }
                }

                return result.ToString();
            }
        }

        public static IEnumerable<object> GetJsonObjectFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var strJson = GetJsonStringFromTabular(fileName, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);

            return JsonConvert.DeserializeObject<IEnumerable<object>>(strJson);
        }

        public static string GetJsonClassModel(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null)
        {
            string result = null;
            var jsonContent = GetJsonStringFromTabular(fileName, skipRows, replaceFrom, replaceTo, headerColumns, true);
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

        #region DataTable results support

        public static DataTable GetDataTable(string fileName, int skipRows = 0, bool useHeader = true)
        {
            DataTable result = null;

            if (!string.IsNullOrWhiteSpace(fileName))
            {
                using (var fileContent = GetFileStream(fileName))
                {
                    result = GetDataTable(fileContent, skipRows, useHeader);
                }
            }

            return result;
        }

        public static DataTable GetDataTable(Stream fileContent, int skipRows = 0, bool useHeader = true)
        {
            DataTable result = null;

            if (fileContent != null)
            {
                var reader = ExcelReaderFactory.CreateReader(fileContent);

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

                result = reader.AsDataSet(config).Tables[0];
            }

            return result;
        }

        #endregion

        #endregion

        #region Form Sheet Parser Public Methods

        public static string GetJsonStringFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("File name not informed");

            using (var engine = new XLWorkbook(fileName))
            {
                var parsedData = ParseFormSheet(engine, sheetName, fieldNames);
                return WriteJsonBodyFromNamedFields(parsedData);
            }
        }

        public static string GetJsonStringFromForm(Stream fileContent, string sheetName, string[] fieldNames, string[] replaceFrom = null, string[] replaceTo = null)
        {
            if (fileContent == null)
                throw new Exception("File content not informed");

            using (var engine = new XLWorkbook(fileContent))
            {
                var parsedData = ParseFormSheet(engine, sheetName, fieldNames);
                return WriteJsonBodyFromNamedFields(parsedData);
            }
        }

        public static object GetJsonObjectFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            var strJson = GetJsonStringFromForm(fileName, sheetName, replaceTo, fieldNames);

            return JsonConvert.DeserializeObject(strJson);
        }

        public static object GetJsonObjectFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            var strJson = GetJsonStringFromForm(fileContent, sheetName, replaceTo, fieldNames);

            return JsonConvert.DeserializeObject(strJson);
        }

        #region Dictionary results support

        public static IDictionary<string, object> GetDictionary(string fileName, string sheetName, string[] fieldNames, string[] replaceFrom = null, string[] replaceTo = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("File name not informed");

            using (var engine = new XLWorkbook(fileName))
            {
                return ParseFormSheet(engine, sheetName, fieldNames);
            }
        }

        public static IDictionary<string, object> GetDictionary(Stream fileContent, string sheetName, string[] fieldNames, string[] replaceFrom = null, string[] replaceTo = null)
        {
            if (fileContent == null)
                throw new Exception("File content not informed");

            using (var engine = new XLWorkbook(fileContent))
            {
                return ParseFormSheet(engine, sheetName, fieldNames);
            }
        }

        #endregion

        #endregion

        #region Tabular Sheet Parser Helper Methods

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
                workbook.SaveToStream(result, FileFormat.Version2010);

                return result;
            }
        }

        private static string[] GetHeaderColumns(IExcelDataReader reader)
        {
            string[] result = null;

            if (reader != null)
            {
                result = new string[reader.FieldCount];

                for (var count = 0; count < reader.FieldCount; count++)
                    result[count] = reader[count]?.ToString().Trim();
            }

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

        private static void WriteItemJsonBodyFromReader(IExcelDataReader reader, JsonWriter writer, string[] headerColumns)
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

        #region Form Sheet Parser Helper Methods

        private static string[] GetNamedFormFields(XLWorkbook engine)
        {
            string[] result = null;

            if (engine != null)
                return engine.NamedRanges.Select(nmf => nmf.Name).ToArray();

            return result;
        }

        private static IDictionary<string, object> GetNamedFieldValues(XLWorkbook engine, string sheetName, string[] fieldNames)
        {
            IDictionary<string, object> result = null;

            result = new Dictionary<string, object>();
            foreach (var field in fieldNames)
            {
                var cell = engine.Cell(field);

                if (cell.Worksheet.Name.ToLower().Equals(sheetName.ToLower()))
                    result.Add(field, cell.Value);
            }

            return result;
        }

        private static string WriteJsonBodyFromNamedFields(IDictionary<string, object> fields)
        {
            using (var result = new StringWriter())
            {
                using (var writer = new JsonTextWriter(result))
                {
                    writer.Formatting = Formatting.Indented;

                    if ((fields != null) && (writer != null))
                    {
                        writer.WriteStartObject();

                        foreach (var field in fields)
                        {
                            writer.WritePropertyName(field.Key);

                            if (field.Value != null)
                                writer.WriteValue(field.Value);
                        }

                        writer.WriteEndObject();
                    }

                    return result.ToString();
                }
            }
        }

        private static IDictionary<string, object> ParseFormSheet(XLWorkbook engine, string sheetName, string[] fieldNames = null)
        {
            if (fieldNames == null)
                fieldNames = GetNamedFormFields(engine);

            return GetNamedFieldValues(engine, sheetName, fieldNames);
        }

        #endregion
    }
}
