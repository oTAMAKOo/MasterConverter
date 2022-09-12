﻿
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Style;

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
                var worksheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var dimension = worksheet.Dimension;

                // セルサイズ調整.
                worksheet.Cells.AutoFitColumns();

                // エラー無視.
                var excelIgnoredError = worksheet.IgnoredErrors.Add(dimension);

                excelIgnoredError.NumberStoredAsText = true;

                // フィールド名取得.
                var fieldNames = ExcelUtility.GetRowValueTexts(worksheet, fieldNameRow)
                    .Select(x => x == null ? null : x.ToLower())
                    .ToArray();

                var fieldIndexDictionary = new Dictionary<string, int>();
                var properties = serializeClass.Class.Properties;

                foreach (var property in properties)
                {
                    int? index = null;

                    var name = property.Key.ToLower();

                    if (fieldNames.Contains(name))
                    {
                        index = Array.FindIndex(fieldNames, x => x.ToLower() == name);
                    }

                    if (!index.HasValue) { continue; }

                    fieldIndexDictionary.Add(name, index.Value);
                }

                // レコード投入用セルを用意.
                for (var i = 0; i < records.Length; i++)
                {
                    var recordRow = recordStartRow + i;

                    // 行追加.
                    if (worksheet.Cells.End.Row < recordRow)
                    {
                        worksheet.InsertRow(recordRow, 1);
                    }

                    // セル情報コピー.
                    for (var column = 1; column < dimension.End.Column; column++)
                    {
                        CloneCellFormat(worksheet, recordStartRow, recordRow, column);
                    }
                }

                // レコード順番入れ替え.

                if (indexData != null && indexData.records != null)
                {
                    var list = new List<Tuple<int, RecordData>>();

                    foreach (var record in records)
                    {
                        var index = indexData.records.IndexOf(x => x == record.recordName);

                        list.Add(Tuple.Create(index, record));
                    }

                    records = list.OrderBy(x => x.Item1).Select(x => x.Item2).ToArray();
                }

                // レコード情報をセルに入力.
                for (var i = 0; i < records.Length; i++)
                {
                    var recordRow = recordStartRow + i;

                    foreach (var recordValue in records[i].values)
                    {
                        var fieldName = recordValue.fieldName.ToLower();
                        var fieldColumn = fieldIndexDictionary.GetValueOrDefault(fieldName, -1);
                        
                        if (fieldColumn == -1) { continue; }

                        object value = null;

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

                                value = string.Format("[{0}]", valueTexts.Any() ? string.Join(",", valueTexts) : string.Empty);
                            }
                            else
                            {
                                value = recordValue.value;
                            }
                        }
                        else
                        {
                            var property = properties.FirstOrDefault(x => x.Key.ToLower() == fieldName);

                            if (!property.Equals(default(Dictionary<string, Type>)))
                            {
                                // Null許容型.
                                if (Nullable.GetUnderlyingType(property.Value) != null)
                                {
                                    value = "null";
                                }
                                else
                                {
                                    value = property.Value.GetDefaultValue();
                                }
                            }
                        }

                        // Excelのセルは1開始なので1加算.
                        var cell = worksheet.Cells[recordRow, fieldColumn + 1];

                        // 折りたたんで全体表示無効.
                        cell.Style.WrapText = false;

                        // 値設定.
                        cell.Value = value;
                    }

                    // セル情報設定.

                    if (records[i].cells != null)
                    {
                        foreach (var cellData in records[i].cells)
                        {
                            ExcelCellUtility.Set<ExcelCell>(worksheet, recordRow, cellData.column, cellData);
                        }
                    }
                }

                // 範囲更新.
                dimension = worksheet.Dimension;
                
                // セルサイズを調整.
                
                var celFitRange = worksheet.Cells[1, 1, dimension.End.Row, dimension.End.Column];

                Func<int, int, string, bool> wrapTextCallback = (r, c, text) =>
                {
                    // 改行が含まれている.
                    return text.FixLineEnd().Contains("\n");
                };

                ExcelUtility.FitColumnSize(worksheet, celFitRange, null, 75, wrapTextCallback);

                ExcelUtility.FitRowSize(worksheet, celFitRange);

                // 保存.
                excel.Save();
            }
        }

        private static void CloneCellFormat(ExcelWorksheet sheet, int recordStartRow, int row, int column)
        {
            var srcCell = sheet.Cells[recordStartRow, column];
            var destCell = sheet.Cells[row, column];

            srcCell.Copy(destCell);
        }
    }
}
