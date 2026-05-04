
using System.Collections;
using System.Globalization;
using Extensions;
using OfficeOpenXml;

namespace MasterConverter
{
    public class EditXlsxBuilder
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void Build(string originXlsxFilePath, SerializeClass serializeClass, IndexData indexData, RecordData[] records, int fieldNameRow, int recordStartRow)
        {
            var masterFolderName = Path.GetFileName(Path.GetDirectoryName(originXlsxFilePath));

            var masterFileName = Path.ChangeExtension(masterFolderName, Constants.MasterFileExtension);

            var editXlsxFilePath = PathUtility.Combine(Path.GetDirectoryName(originXlsxFilePath), masterFileName);

            // ファイルが存在＋ロック時はエラー.
            if (File.Exists(editXlsxFilePath))
            {
                if (FileUtility.IsFileLocked(editXlsxFilePath))
                {
                    throw new FileLoadException(string.Format("File locked!!\n\n{0}", editXlsxFilePath));
                }
            }

            //------ エディット用にxlsxを複製 ------

            var originXlsxFile = new FileInfo(originXlsxFilePath);

            var editXlsxFile = originXlsxFile.CopyTo(editXlsxFilePath, true);
            
            //------ レコード情報を書き込み ------

            using (var excel = new ExcelPackage(editXlsxFile))
            {
                var workBook = excel.Workbook;
                var worksheet = workBook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var dimension = worksheet.Dimension;

                // カラム初期幅.

                var columnsWidth = new Dictionary<int, double>();

                for (var c = dimension.Start.Column; c <= dimension.End.Column; c++)
                {
                    var width = worksheet.Columns[c].Width;

                    columnsWidth[c] = width;
                }

                // エラー無視.
                var excelIgnoredError = worksheet.IgnoredErrors.Add(dimension);

                excelIgnoredError.NumberStoredAsText = true;

                // フィールド名取得.
                var fieldNames = ExcelUtility.GetRowValueTexts(worksheet, fieldNameRow)
                    .Select(x => x == null ? null : x.ToLower())
                    .ToArray();

                var properties = serializeClass.Class.Properties;

                // フィールド名 → カラムインデックスの逆引き辞書.
                var fieldNameToIndex = new Dictionary<string, int>();

                for (var i = 0; i < fieldNames.Length; i++)
                {
                    if (fieldNames[i] != null && !fieldNameToIndex.ContainsKey(fieldNames[i]))
                    {
                        fieldNameToIndex[fieldNames[i]] = i;
                    }
                }

                var fieldIndexDictionary = new Dictionary<string, int>();

                foreach (var property in properties)
                {
                    var name = property.Key.ToLower();

                    if (fieldNameToIndex.TryGetValue(name, out var idx))
                    {
                        fieldIndexDictionary[name] = idx;
                    }
                }

                // properties の大文字小文字無視辞書（null値のデフォルト値取得用）.
                var propertiesLookup = new Dictionary<string, KeyValuePair<string, Type>>(StringComparer.OrdinalIgnoreCase);

                foreach (var property in properties)
                {
                    if (!propertiesLookup.ContainsKey(property.Key))
                    {
                        propertiesLookup[property.Key] = property;
                    }
                }

                // フィールド名 ToLower キャッシュ.
                var lowerFieldNameCache = new Dictionary<string, string>();

                // レコード投入用セルを一括挿入（テンプレート行のスタイルをコピー）.
                if (records.Length > 1)
                {
                    worksheet.InsertRow(recordStartRow + 1, records.Length - 1, recordStartRow);
                }

                // レコード順番入れ替え.

                if (indexData != null && indexData.records != null)
                {
                    var indexLookup = new Dictionary<string, int>(indexData.records.Length);

                    for (var i = 0; i < indexData.records.Length; i++)
                    {
                        if (!indexLookup.ContainsKey(indexData.records[i]))
                        {
                            indexLookup[indexData.records[i]] = i;
                        }
                    }

                    var list = new List<Tuple<int, RecordData>>(records.Length);

                    foreach (var record in records)
                    {
                        var index = indexLookup.TryGetValue(record.recordName, out var idx) ? idx : -1;

                        list.Add(Tuple.Create(index, record));
                    }

                    records = list.OrderBy(x => x.Item1).Select(x => x.Item2).ToArray();
                }

                // 各列の最長値をレコードデータから事前収集（AutoFitColumns用）.
                var maxCellValues = new object[dimension.End.Column + 1];
                var maxCellLengths = new int[dimension.End.Column + 1];

                foreach (var record in records)
                {
                    foreach (var recordValue in record.values)
                    {
                        if (!lowerFieldNameCache.TryGetValue(recordValue.fieldName, out var fn))
                        {
                            fn = recordValue.fieldName.ToLower();
                            lowerFieldNameCache[recordValue.fieldName] = fn;
                        }

                        var col = fieldIndexDictionary.GetValueOrDefault(fn, -1);

                        if (col == -1) { continue; }

                        var cellValue = ConvertRecordValue(recordValue, fn, propertiesLookup);
                        var valueText = cellValue != null ? cellValue.ToString() : string.Empty;
                        var len = valueText.Length;

                        if (len > maxCellLengths[col])
                        {
                            maxCellLengths[col] = len;
                            maxCellValues[col] = cellValue;
                        }
                    }
                }

                // 末尾に仮行を追加して最長値を書き込み.
                var tempRow = dimension.End.Row + 1;

                for (var c = dimension.Start.Column; c <= dimension.End.Column; c++)
                {
                    if (maxCellValues[c] != null)
                    {
                        worksheet.Cells[tempRow, c].Value = maxCellValues[c];
                    }
                }

                // ヘッダー行 + 仮行でAutoFitColumns.
                var fitRange = worksheet.Cells[1, dimension.Start.Column, tempRow, dimension.End.Column];
                fitRange.AutoFitColumns();

                // 仮行を削除.
                worksheet.DeleteRow(tempRow);

                // レコード情報をセルに入力.
                for (var i = 0; i < records.Length; i++)
                {
                    var recordRow = recordStartRow + i;

                    foreach (var recordValue in records[i].values)
                    {
                        if (!lowerFieldNameCache.TryGetValue(recordValue.fieldName, out var fieldName))
                        {
                            fieldName = recordValue.fieldName.ToLower();
                            lowerFieldNameCache[recordValue.fieldName] = fieldName;
                        }

                        var fieldColumn = fieldIndexDictionary.GetValueOrDefault(fieldName, -1);

                        if (fieldColumn == -1) { continue; }

                        // Excelのセルは1開始なので1加算.
                        var cell = worksheet.Cells[recordRow, fieldColumn + 1];

                        // 数式があるセルへはデータ入力を行わない.
                        if(!string.IsNullOrEmpty(cell.Formula)) { continue; }

                        var value = ConvertRecordValue(recordValue, fieldName, propertiesLookup);

                        // 折り畳んで全体表示無効.
                        var wrapText = false;

                        // 改行を含む場合は折り畳む.
                        if (value is string text && !string.IsNullOrEmpty(text))
                        {
                            wrapText = text.FixLineEnd().Contains("\n");
                        }

                        cell.Style.WrapText = wrapText;

                        // 値設定.
                        cell.Value = value;
                    }

                    // セル情報設定.

                    if (records[i].cells != null)
                    {
                        foreach (var cellData in records[i].cells)
                        {
                            ExcelCellUtility.Set(worksheet, recordRow, cellData.column, cellData);
                        }
                    }
                }

                // 範囲更新.
                dimension = worksheet.Dimension;

                for (var c = dimension.Start.Column; c <= dimension.End.Column; c++)
                {
                    var baseWidth = columnsWidth.GetValueOrDefault(c);
                    var currentWidth = worksheet.Column(c).Width;

                    if (currentWidth < baseWidth)
                    {
                        worksheet.Column(c).Width = baseWidth;
                    }
                }

                // マスターのシートをアクティブにする.

                var avtiveTab = workBook.Worksheets.IndexOf(x => x.Name == Constants.MasterSheetName);

                workBook.View.ActiveTab = avtiveTab;
                worksheet.Select();

                // 保存.
                excel.Save();
            }
        }

        private static object ConvertRecordValue(RecordValue recordValue, string fieldNameLower, Dictionary<string, KeyValuePair<string, Type>> propertiesLookup)
        {
            if (recordValue.value != null)
            {
                var type = recordValue.value.GetType();

                // 配列.
                if (type.IsArray)
                {
                    var array = recordValue.value as IEnumerable;

                    var valueTexts = new List<string>();

                    foreach (var item in array)
                    {
                        valueTexts.Add(item.ToString());
                    }

                    return string.Format("[{0}]", valueTexts.Any() ? string.Join(",", valueTexts) : string.Empty);
                }
                else if (type == typeof(DateTime))
                {
                    return recordValue.value.ToString();
                }
                else if (type == typeof(float))
                {
                    // float→doubleでIEEE 754の精度誤差が露出するため文字列経由で変換.
                    return double.Parse(((float)recordValue.value).ToString("R", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
                }
                else
                {
                    return recordValue.value;
                }
            }
            else
            {
                if (propertiesLookup.TryGetValue(fieldNameLower, out var property))
                {
                    // Null許容型.
                    if (Nullable.GetUnderlyingType(property.Value) != null)
                    {
                        return "null";
                    }
                    else
                    {
                        return property.Value.GetDefaultValue();
                    }
                }
            }

            return null;
        }
    }
}
