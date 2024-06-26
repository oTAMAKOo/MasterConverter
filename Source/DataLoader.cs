﻿
using System.Collections.Generic;
using System.Text;
using Extensions;
using Newtonsoft.Json;
using OfficeOpenXml;
using YamlDotNet.Serialization;

namespace MasterConverter
{
    public static class DataLoader
    {
        //----- params -----

        //----- field -----

        private static JsonSerializerSettings jsonSerializerSettings = null;

        private static IDeserializer yamlDeserializer = null;

        //----- property -----

        //----- method -----

        /// <summary> Importが必要か </summary>
        public static bool IsRequireImport(string excelFilePath, string yamlDirectory)
        {
            if (!File.Exists(excelFilePath)){ return true; }

            var lastUpdateTime = File.GetLastWriteTimeUtc(excelFilePath);

            // IndexFile.
            {
                var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

                if (File.Exists(filePath))
                {
                    if(lastUpdateTime < File.GetLastWriteTimeUtc(filePath)){ return true; }
                }
            }

            // RecordFile.
            {
                if (Directory.Exists(yamlDirectory))
                {
                    var filePaths = Directory.EnumerateFiles(yamlDirectory, "*.*")
                        .Where(x =>
                           {
                               var extension = Path.GetExtension(x);

                               return extension == Constants.RecordFileExtension || extension == Constants.CellOptionFileExtension;
                           })
                        .ToArray();

                    foreach (var filePath in filePaths)
                    {
                        if (lastUpdateTime < File.GetLastWriteTimeUtc(filePath)){ return true; }
                    }
                }
            }

            return false;
        }

        /// <summary> Exportが必要か </summary>
        public static bool IsRequireExport(string excelFilePath, string yamlDirectory)
        {
            if (!File.Exists(excelFilePath)){ return false; }

            var lastUpdateTime = File.GetLastWriteTimeUtc(excelFilePath);

            // IndexFile.
            {
                var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

                if (File.Exists(filePath))
                {
                    if(File.GetLastWriteTimeUtc(filePath) < lastUpdateTime){ return true; }
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        public static IndexData LoadRecordIndex(string excelFilePath, SerializationFileUtility.Format format)
        {
            var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

            return SerializationFileUtility.LoadFile<IndexData>(filePath, format);
        }

        /// <summary> レコード情報読み込み(.yaml) </summary>
        public static async Task<RecordData[]> LoadRecords(string recordDirectory, TypeGenerator typeGenerator, SerializationFileUtility.Format format)
        {
            if (!Directory.Exists(recordDirectory)) { return new RecordData[0]; }

            var recordFiles = Directory.EnumerateFiles(recordDirectory, "*.*")
                .Where(x => Path.GetExtension(x) == Constants.RecordFileExtension)
                .OrderBy(x => x, new NaturalComparer())
                .ToArray();

            // レコードファイル読み込み.

            var recordData = new Dictionary<string, string>();

            if (recordFiles.Any())
            {
                var tasks = new List<Task>();

                foreach (var recordFile in recordFiles)
                {
                    var task = Task.Run(async () =>
                    {
                        using (var file = new FileStream(recordFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            using (var reader = new StreamReader(file, Encoding.UTF8))
                            {
                                var text = await reader.ReadToEndAsync();

                                lock (recordData)
                                {
                                    recordData.Add(recordFile, text);
                                }
                            }
                        }
                    });

                    tasks.Add(task);
                }

                await Task.WhenAll(tasks);
            }

            // オプションデータ読み込み.

            var optionDataDictionary = new Dictionary<string, ExcelCell[]>();

            if (recordData.Any())
            {
                var tasks = new List<Task>();

                foreach (var item in recordData)
                {
                    var recordFilePath = item.Key;

                    var optionFilePath = Path.ChangeExtension(recordFilePath, Constants.CellOptionFileExtension);

                    if (!File.Exists(optionFilePath)) { continue; }

                    var task = Task.Run(() =>
                    {
                        var optionData = SerializationFileUtility.LoadFile<ExcelCell[]>(optionFilePath, format);

                        lock (optionDataDictionary)
                        {
                            if (!optionDataDictionary.ContainsKey(optionFilePath))
                            {
                                optionDataDictionary.Add(optionFilePath, new ExcelCell[0]);
                            }

                            optionDataDictionary[recordFilePath] = optionData;
                        }
                    });

                    tasks.Add(task);
                }

                await Task.WhenAll(tasks);
            }

            // レコードクラスに変換.

            var recordList = new List<RecordData>();

            if (recordData.Any())
            {
                var tasks = new List<Task>();
                
                foreach (var item in recordData)
                {
                    var task = Task.Run(() =>
                    {
                        var instance = Deserialize(typeGenerator.Type, item.Value, format);

                        var recordValues = new List<RecordValue>();

                        foreach (var property in typeGenerator.Properties)
                        {
                            var value = TypeGenerator.GetProperty(instance, property.Key, property.Value);

                            var recordValue = new RecordValue()
                            {
                                fieldName = property.Key,
                                value = value,
                            };

                            recordValues.Add(recordValue);
                        }

                        var values = recordValues.ToArray();

                        var record = new RecordData()
                        {
                            recordName = GetRecordName(string.Empty, values),
                            values = values,
                            cells = optionDataDictionary.GetValueOrDefault(item.Key),
                        };

                        lock (recordList)
                        {
                            recordList.Add(record);
                        }
                    });

                    tasks.Add(task);
                }

                await Task.WhenAll(tasks);
            }

            // レコード名を重複しない形式に更新.
            recordList = UpdateRecordNames(recordList);

            return recordList.ToArray();
        }

        private static object Deserialize(Type type, string text, SerializationFileUtility.Format format)
        {
            object result = null;

            switch (format)
            {
                case SerializationFileUtility.Format.Json:
                {
                    if (jsonSerializerSettings == null)
                    {
                        jsonSerializerSettings = new JsonSerializerSettings()
                        {
                            Formatting = Formatting.Indented,
                            NullValueHandling = NullValueHandling.Ignore,
                        };
                    }

                    result = JsonConvert.DeserializeObject(text, type, jsonSerializerSettings);
                }
                    break;

                case SerializationFileUtility.Format.Yaml:
                {
                    if (yamlDeserializer == null)
                    {
                        var builder = new DeserializerBuilder();

                        builder.IgnoreUnmatchedProperties();

                        yamlDeserializer = builder.Build();
                    }

                    result = yamlDeserializer.Deserialize(text, type);
                }
                    break;
            }

            return result;
        }

        /// <summary> レコード情報読み込み(.xlsx) </summary>
        public static RecordData[] LoadExcelRecords(string excelFilePath, int fieldNameRow, int recordStartRow)
        {
            var recordList = new List<RecordData>();

            if (!File.Exists(excelFilePath)) { return new RecordData[0]; }

            using (var excel = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var address = worksheet.Dimension;

                var fieldNames = ExcelUtility.GetRowValueTexts(worksheet, fieldNameRow).ToArray();
                
                for (var r = recordStartRow; r <= address.End.Row; r++)
                {
                    var recordValues = new List<RecordValue>();

                    var records = ExcelUtility.GetRowValues(worksheet, r).ToArray();

                    for (var i = 0; i < records.Length; i++)
                    {
                        var fieldName = fieldNames[i];

                        if (string.IsNullOrEmpty(fieldName)) { continue; }

                        var recordValue = new RecordValue()
                        {
                            fieldName = fieldName,
                            value = records[i],
                        };

                        recordValues.Add(recordValue);
                    }

                    // セル情報取得.

                    var cells = new List<ExcelCell>();

                    for (var c = 1; c < records.Length + 1; c++)
                    {
                        var cellData = ExcelCellUtility.Get<ExcelCell>(worksheet, r, c);

                        if (cellData == null) { continue; }

                        cellData.column = c;

                        cells.Add(cellData);
                    }

                    var values = recordValues.ToArray();

                    var record = new RecordData()
                    {
                        recordName = GetRecordName(string.Empty, values),
                        values = values,
                        cells = cells.Any() ? cells.ToArray() : null,
                    };

                    recordList.Add(record);
                }
            }

            // 全レコード重複のデータが存在するか.
            var duplicates =HasDuplication(recordList);

            if (duplicates.Any())
            {
                var builder = new StringBuilder();

                duplicates.ForEach(x => builder.AppendLine(x));

                throw new Exception($"Duplication records exist!\nFile : {excelFilePath}\n{builder}");
            }

            // レコード名を重複しない形式に更新.
            recordList = UpdateRecordNames(recordList);

            return recordList.ToArray();
        }

        /// <summary> 全レコード重複のデータが存在するか </summary>
        public static string[] HasDuplication(List<RecordData> records)
        {
            var duplicates = records
                .Where(x => x.values.Any(y => y.value != null))
                .Select(x => String.Join(",", x.values.Select(y => y.value)))
                .GroupBy(name => name)
                .Where(name => name.Count() > 1)
                .Select(group => group.Key)
                .ToArray();

            return duplicates;
        }

        private static List<RecordData> UpdateRecordNames(List<RecordData> records)
        {
            while (true)
            {
                var rename = false;

                var groups = records.Where(x => !string.IsNullOrEmpty(x.recordName))
                    .GroupBy(x => x.recordName)
                    .ToArray();

                foreach (var group in groups)
                {
                    if (1 < group.Count())
                    {
                        foreach (var item in group)
                        {
                            item.recordName = GetRecordName(item.recordName, item.values);
                        }

                        rename = true;
                    }
                }

                records = groups.SelectMany(x => x).ToList();

                if (!rename) { break; }
            }

            return records;
        }

        private static string GetRecordName(string name, RecordValue[] records)
        {
            var index = 0;
            var newName = string.Empty;

            while (name.StartsWith(newName))
            {
                if (records.Length <= index) { break; }

                var recordValue = records[index++];

                if (recordValue == null) { continue; }

                if (recordValue.fieldName.StartsWith(Constants.IgnoreFieldPrefix)){ continue; }

                var value = recordValue.value != null ? recordValue.value.ToString() : null;

                if (string.IsNullOrEmpty(value)) { continue; }

                newName = string.IsNullOrEmpty(newName) ? value : string.Format("{0}_{1}", newName, value);
            }

            return newName;
        }
    }
}
