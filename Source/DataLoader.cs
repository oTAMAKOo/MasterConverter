
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

        private static readonly Lazy<JsonSerializerSettings> jsonSerializerSettings = new Lazy<JsonSerializerSettings>(() =>
            new JsonSerializerSettings()
            {
                Formatting = Formatting.Indented,
                NullValueHandling = NullValueHandling.Ignore,
            });

        private static readonly Lazy<IDeserializer> yamlDeserializer = new Lazy<IDeserializer>(() =>
        {
            var builder = new DeserializerBuilder();
            builder.IgnoreUnmatchedProperties();
            return builder.Build();
        });

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

            var count = recordFiles.Length;

            if (count == 0) { return new RecordData[0]; }

            // 事前確保配列（インデックス書き込みでロック不要）.
            var fileTexts = new string[count];
            var optionDataArray = new ExcelCell[count][];
            var recordArray = new RecordData[count];

            // レコードファイル読み込み.

            var readTasks = new Task[count];

            for (var i = 0; i < count; i++)
            {
                var idx = i;

                readTasks[idx] = Task.Run(async () =>
                {
                    using (var file = new FileStream(recordFiles[idx], FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (var reader = new StreamReader(file, Encoding.UTF8))
                        {
                            fileTexts[idx] = await reader.ReadToEndAsync();
                        }
                    }
                });
            }

            await Task.WhenAll(readTasks);

            // オプションデータ読み込み.

            var optionTasks = new Task[count];

            for (var i = 0; i < count; i++)
            {
                var idx = i;
                var optionFilePath = Path.ChangeExtension(recordFiles[idx], Constants.CellOptionFileExtension);

                if (File.Exists(optionFilePath))
                {
                    optionTasks[idx] = Task.Run(() =>
                    {
                        optionDataArray[idx] = SerializationFileUtility.LoadFile<ExcelCell[]>(optionFilePath, format);
                    });
                }
                else
                {
                    optionTasks[idx] = Task.CompletedTask;
                }
            }

            await Task.WhenAll(optionTasks);

            // レコードクラスに変換.

            var deserializeTasks = new Task[count];

            for (var i = 0; i < count; i++)
            {
                var idx = i;

                deserializeTasks[idx] = Task.Run(() =>
                {
                    var instance = Deserialize(typeGenerator.Type, fileTexts[idx], format);

                    var recordValues = new List<RecordValue>();

                    foreach (var property in typeGenerator.Properties)
                    {
                        var value = typeGenerator.GetCachedProperty(instance, property.Key);

                        var recordValue = new RecordValue()
                        {
                            fieldName = property.Key,
                            value = value,
                        };

                        recordValues.Add(recordValue);
                    }

                    var values = recordValues.ToArray();

                    recordArray[idx] = new RecordData()
                    {
                        recordName = GetRecordName(string.Empty, values),
                        values = values,
                        cells = optionDataArray[idx],
                    };
                });
            }

            await Task.WhenAll(deserializeTasks);

            // レコード名を重複しない形式に更新.
            var recordList = UpdateRecordNames(recordArray.ToList());

            return recordList.ToArray();
        }

        private static object Deserialize(Type type, string text, SerializationFileUtility.Format format)
        {
            object result = null;

            switch (format)
            {
                case SerializationFileUtility.Format.Json:
                {
                    result = JsonConvert.DeserializeObject(text, type, jsonSerializerSettings.Value);
                }
                    break;

                case SerializationFileUtility.Format.Yaml:
                {
                    result = yamlDeserializer.Value.Deserialize(text, type);
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

            var rowTracker = new Dictionary<string, List<(int row, RecordValue[] values)>>();

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

                    // 全フィールドが空のレコードはスキップ.
                    if (values.All(v => string.IsNullOrEmpty(v.value?.ToString()))) { continue; }

                    var record = new RecordData()
                    {
                        recordName = GetRecordName(string.Empty, values),
                        values = values,
                        cells = cells.Any() ? cells.ToArray() : null,
                    };

                    var rowKey = string.Join(",", values.Select(y => y.value));
                    if (!rowTracker.ContainsKey(rowKey))
                        rowTracker[rowKey] = new List<(int, RecordValue[])>();
                    rowTracker[rowKey].Add((r, values));

                    recordList.Add(record);
                }
            }

            // 全レコード重複のデータが存在するか.
            var duplicateGroups = rowTracker
                .Where(x => x.Value.Count > 1 && x.Value[0].values.Any(y => y.value != null))
                .ToArray();

            if (duplicateGroups.Any())
            {
                var builder = new StringBuilder();

                foreach (var dup in duplicateGroups)
                {
                    var rows = string.Join(", ", dup.Value.Select(x => x.row));
                    var nonEmptyFields = dup.Value[0].values.Where(v => v.value != null).ToArray();
                    var fields = nonEmptyFields.Any()
                        ? string.Join(", ", nonEmptyFields.Select(v => $"{v.fieldName}={v.value}"))
                        : "(empty)";
                    builder.AppendLine($"  [Row {rows}] {fields}");
                }

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
            // 各レコードの有効フィールド値を事前計算.
            var entries = new List<(RecordData record, string[] fieldValues)>(records.Count);

            foreach (var record in records)
            {
                if (string.IsNullOrEmpty(record.recordName))
                {
                    entries.Add((record, Array.Empty<string>()));
                    continue;
                }

                var values = record.values
                    .Where(v => v != null
                        && !v.fieldName.StartsWith(Constants.IgnoreFieldPrefix)
                        && !string.IsNullOrEmpty(v.value?.ToString()))
                    .Select(v => v.value.ToString())
                    .ToArray();

                entries.Add((record, values));
            }

            var maxDepth = entries.Count > 0
                ? entries.Max(x => x.fieldValues.Length)
                : 0;

            var currentNames = new string[entries.Count];
            var finalized = new bool[entries.Count];

            for (var depth = 1; depth <= maxDepth; depth++)
            {
                // 現在の深さまでの名前を構築.
                for (var i = 0; i < entries.Count; i++)
                {
                    if (finalized[i]) { continue; }

                    var fields = entries[i].fieldValues;

                    currentNames[i] = string.Join("_", fields.Take(depth));
                }

                // Dictionaryで重複カウント.
                var nameCount = new Dictionary<string, int>();

                for (var i = 0; i < entries.Count; i++)
                {
                    if (finalized[i]) { continue; }

                    nameCount.TryGetValue(currentNames[i], out var count);
                    nameCount[currentNames[i]] = count + 1;
                }

                // 重複なしのレコードを確定.
                var allResolved = true;

                for (var i = 0; i < entries.Count; i++)
                {
                    if (finalized[i]) { continue; }

                    if (nameCount[currentNames[i]] == 1)
                    {
                        finalized[i] = true;
                        entries[i].record.recordName = currentNames[i];
                    }
                    else
                    {
                        allResolved = false;
                    }
                }

                if (allResolved) { break; }
            }

            // 最大深さでも解決できなかったレコード.
            for (var i = 0; i < entries.Count; i++)
            {
                if (!finalized[i])
                {
                    entries[i].record.recordName = currentNames[i];
                }
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
