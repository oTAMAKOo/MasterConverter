
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using YamlDotNet.Serialization;

namespace MasterConverter
{
    public static class DataLoader
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static IndexData LoadRecordIndex(string excelFilePath)
        {
            var filePath = Path.ChangeExtension(excelFilePath, Constants.IndexFileExtension);

            return FileSystem.LoadFile<IndexData>(filePath, FileSystem.Format.Yaml);
        }

        /// <summary> レコード情報読み込み(.yaml) </summary>
        public static RecordData[] LoadYamlRecords(string yamlDirectory, TypeGenerator typeGenerator)
        {
            if (!Directory.Exists(yamlDirectory)) { return new RecordData[0]; }

            var recordFiles = Directory.EnumerateFiles(yamlDirectory, "*.*")
                .Where(x => Path.GetExtension(x) == Constants.RecordFileExtension)
                .OrderBy(x => x, new NaturalComparer())
                .ToArray();

            var recordData = new Dictionary<string, string>();

            foreach (var recordFile in recordFiles)
            {
                using (var file = new FileStream(recordFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = new StreamReader(file, Encoding.UTF8))
                    {
                        recordData.Add(recordFile, reader.ReadToEnd());
                    }
                }
            }

            var optionDataDictionary = new Dictionary<string, ExcelCell[]>();

            foreach (var item in recordData)
            {
                var recordFilePath = item.Key;

                var optionFilePath = Path.ChangeExtension(recordFilePath, Constants.CellOptionFileExtension);

                if (!File.Exists(optionFilePath)) { continue; }

                if(!optionDataDictionary.ContainsKey(optionFilePath))
                {
                    optionDataDictionary.Add(optionFilePath, new ExcelCell[0]);
                }

                var optionData = FileSystem.LoadFile<ExcelCell[]>(optionFilePath, FileSystem.Format.Yaml);

                optionDataDictionary[recordFilePath] = optionData;
            }

            var recordList = new List<RecordData>();
            
            var deserializer = new DeserializerBuilder().IgnoreUnmatchedProperties().Build();

            foreach (var item in recordData)
            {
                var instance = deserializer.Deserialize(item.Value, typeGenerator.Type);

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

                recordList.Add(record);
            };

            // レコード名を重複しない形式に更新.
            recordList = UpdateRecordNames(recordList);

            return recordList.ToArray();
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

            // レコード名を重複しない形式に更新.
            recordList = UpdateRecordNames(recordList);

            return recordList.ToArray();
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

                var value = recordValue.value != null ? recordValue.value.ToString() : null;

                if (string.IsNullOrEmpty(value)) { continue; }

                newName = string.IsNullOrEmpty(newName) ? value : string.Format("{0}_{1}", newName, value);
            }

            return newName;
        }
    }
}
