﻿
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Extensions;
using OfficeOpenXml;

namespace MasterConverter
{
    public class SerializeClass
    {
        //----- params -----

        private class Property
        {
            public bool export;
            public Type type;
            public string fieldName;
        }

        //----- field -----

        // プロパティ情報.
        private Property[] properties = null;

        //----- property -----

        public TypeGenerator Class { get; private set; }

        //----- method -----

        /// <summary> Excelからクラス情報を読み込み </summary>
        public void LoadClassSchema(string classSchemaFilePath, string[] exportTags, int tagRow, int dataTypeRow, int fieldNameRow)
        {
            if(!File.Exists(classSchemaFilePath))
            {
                throw new FileNotFoundException(string.Format("ClassSchema file not found. {0}", classSchemaFilePath));
            }

            exportTags = exportTags ?? new string[0];

            using (var excel = new ExcelPackage(new FileInfo(classSchemaFilePath)))
            {
                var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var address = sheet.Dimension;
                
                var tags = ExcelUtility.GetRowValueTexts(sheet, tagRow).ToArray();
                var dataTypes = ExcelUtility.GetRowValueTexts(sheet, dataTypeRow).ToArray();
                var fieldNames = ExcelUtility.GetRowValueTexts(sheet, fieldNameRow).ToArray();

                var list = new List<Property>();

                for (var i = 0; i < address.End.Column; i++)
                {
                    if (dataTypes[i] == null || fieldNames[i] == null) { break; }

                    var property = new Property()
                    {
                        export = exportTags.IsEmpty() || exportTags.Contains(tags[i]),
                        type = TypeUtility.GetTypeFromSystemTypeName(dataTypes[i]),
                        fieldName = fieldNames[i].Trim(' ', '　', '\n', '\t'),
                    };

                    list.Add(property);
                }

                properties = list.ToArray();
            }

            var dictionary = new Dictionary<string, Type>();

            foreach (var info in properties)
            {
                if (!info.export) { continue; }

                if (string.IsNullOrEmpty(info.fieldName)) { continue; }

                if (info.type == null) { continue; }

                dictionary.Add(info.fieldName, info.type);
            }

            Class = new TypeGenerator("SerializeClass", dictionary);
        }

        /// <summary> レコードからクラスを生成 </summary>
        public object CreateInstance(Dictionary<string, string> fieldValues)
        {
            var instance = Class.NewInstance();

            foreach (var item in fieldValues)
            {
                var property = properties.FirstOrDefault(x => x.fieldName == item.Key);

                if (property == null) { continue; }

                if (!property.export) { continue; }

                try
                {
                    var value = ParseValue(item.Value, property.type);

                    TypeGenerator.SetProperty(instance, property.fieldName, value, property.type);
                }
                catch
                {
                    Console.WriteLine(string.Format("CsvRecord error. [ERROR] ({0}){1} = {2}\n", property.type.Name, property.fieldName, item.Value));
                    throw;
                }
            }

            return instance;
        }

        public string GetSchemaString()
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (var property in properties)
            {
                if (property == null) { continue; }
                if (!property.export) { continue; }
                Console.WriteLine(property.fieldName);
                if (property.type == null) { continue; }
                
                stringBuilder.AppendLine(property.fieldName);
                bool IsNullable(Type type) => Nullable.GetUnderlyingType(type) != null;
                if (IsNullable(property.type))
                {
                    stringBuilder.AppendLine("Nullable " + Nullable.GetUnderlyingType(property.type).FullName);
                }
                else
                {
                    stringBuilder.AppendLine(property.type.FullName);
                }
                stringBuilder.AppendLine();
            }
            return stringBuilder.ToString();
        }

        private object ParseValue(string valueText, Type valueType)
        {
            var value = valueType.GetDefaultValue();

            // Null許容型.
            var underlyingType = Nullable.GetUnderlyingType(valueType);

            if (underlyingType != null)
            {
                valueType = underlyingType;

                if (!string.IsNullOrEmpty(valueText))
                {
                    if (valueText.ToLower() == "null")
                    {
                        valueText = string.Empty;
                    }
                }
            }

            // 空文字列ならデフォルト値.
            if (string.IsNullOrEmpty(valueText))
            {
                // Null許容型.
                if (underlyingType != null) { return null; }

                // 配列.
                if (valueType.IsArray)
                {
                    var elementType = valueType.GetElementType();

                    return Array.CreateInstance(elementType, 0);
                }

                return value;
            }

            // 配列.
            if (valueType.IsArray)
            {
                var list = new List<object>();

                var elementType = valueType.GetElementType();

                var arrayText = valueText;

                var start = arrayText.IndexOf("[", StringComparison.Ordinal);
                var end = arrayText.LastIndexOf("]", StringComparison.Ordinal);

                // 複数要素のある場合.
                if (start != -1 && end != -1 && start < end)
                {
                    // 「[]」を外す.
                    arrayText = arrayText.Substring(start + 1, end - start - 1);

                    // 「,」区切りで配列化.
                    var elements = arrayText.Split(',').Where(x => !string.IsNullOrEmpty(x)).ToArray();

                    foreach (var element in elements)
                    {
                        list.Add(ConvertValue(element, elementType));
                    }
                }
                // 「[]」で囲まれてない場合は1つしか要素がない配列に変換.
                else
                {
                    list.Add(ConvertValue(valueText, elementType));
                }

                var array = Array.CreateInstance(elementType, list.Count);

                Array.Copy(list.ToArray(), array, list.Count);

                value = array;
            }
            // 単一要素.
            else
            {
                value = ConvertValue(valueText, valueType);
            }

            return value;
        }

        private static object ConvertValue(string valueText, Type valueType)
        {
            if (valueType != typeof(string))
            {
                valueText = valueText.Trim(' ', '　', '\n', '\t');
            }

            return Convert.ChangeType(valueText, valueType, CultureInfo.InvariantCulture);
        }
    }
}
