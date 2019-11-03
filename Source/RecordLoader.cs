﻿
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
    public static class RecordLoader
    {
        //----- params -----

        public class RecordData
        {
            public int index;
            public string recordName;
            public RecordValue[] values;
        }

        public class RecordValue
        {
            public string fieldName;
            public object value;
        }

        //----- field -----

        //----- property -----

        //----- method -----

        /// <summary> レコード情報読み込み(.yaml) </summary>
        public static RecordData[] LoadYamlRecords(string yamlDirectory, TypeGenerator typeGenerator)
        {
            if (!Directory.Exists(yamlDirectory)) { return new RecordData[0]; }

            var recordFiles = Directory.EnumerateFiles(yamlDirectory, "*.*")
                .Where(x => Path.GetExtension(x) == Constants.RecordFileExtension)
                .OrderBy(x => x, new NaturalComparer())
                .ToArray();

            var list = new List<string>();

            foreach (var recordFile in recordFiles)
            {
                using (var file = new FileStream(recordFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = new StreamReader(file, Encoding.UTF8))
                    {
                        list.Add(reader.ReadToEnd());
                    }
                }
            }

            var recordList = new List<RecordData>();
            
            var deserializer = new DeserializerBuilder().IgnoreUnmatchedProperties().Build();

            for (var i = 0; i < list.Count; i++)
            {
                var instance = deserializer.Deserialize(list[i], typeGenerator.Type);

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
                    index = i,
                    recordName = GetRecordName(string.Empty, values),
                    values = values,
                };

                recordList.Add(record);
            };

            // レコード名を重複しない形式に更新.
            recordList = UpdateRecordNames(recordList);

            return recordList.ToArray();
        }

        /// <summary> レコード情報読み込み(.xlsx) </summary>
        public static RecordData[] LoadXlsxRecords(string xlsxFilePath, int fieldNameRow, int recordStartRow)
        {
            var recordList = new List<RecordData>();

            if (!File.Exists(xlsxFilePath)) { return new RecordData[0]; }

            using (var excel = new ExcelPackage(new FileInfo(xlsxFilePath)))
            {
                var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var address = sheet.Dimension;

                var fieldNames = ExcelUtility.GetRowValueTexts(sheet, fieldNameRow).ToArray();
                
                for (var r = recordStartRow; r <= address.End.Row; r++)
                {
                    var recordValues = new List<RecordValue>();
                    
                    var records = ExcelUtility.GetRowValues(sheet, r).ToArray();

                    for (var c = 0; c < records.Length; c++)
                    {
                        var fieldName = fieldNames[c];

                        if (string.IsNullOrEmpty(fieldName)) { continue; }

                        var recordValue = new RecordValue()
                        {
                            fieldName = fieldName,
                            value = records[c],
                        };

                        recordValues.Add(recordValue);
                    }

                    var values = recordValues.ToArray();

                    var record = new RecordData()
                    {
                        index = r,
                        recordName = GetRecordName(string.Empty, values),
                        values = values,
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
