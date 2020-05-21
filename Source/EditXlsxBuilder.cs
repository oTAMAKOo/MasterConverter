
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

        public static void Build(string originXlsxFilePath, SerializeClass serializeClass, 
                                 RecordLoader.RecordData[] records, CellOptionLoader.CellOption[] cellOptions, 
                                 int fieldNameRow, int recordStartRow)
        {
            var masterFolderName = Path.GetFileName(Path.GetDirectoryName(originXlsxFilePath));

            var masterFileName = Path.ChangeExtension(masterFolderName, Constants.MasterFileExtension);

            var editXlsxFilePath = PathUtility.Combine(Path.GetDirectoryName(originXlsxFilePath), masterFileName);

            // ファイルが存在＋ロック時はエラー.
            if (File.Exists(editXlsxFilePath))
            {
                if (FileUtility.IsFileLocked(editXlsxFilePath))
                {
                    throw new FileLoadException(string.Format("File locked. {0}", editXlsxFilePath));
                }
            }

            //------ エディット用にxlsxを複製 ------

            var originXlsxFile = new FileInfo(originXlsxFilePath);

            var editXlsxFile = originXlsxFile.CopyTo(editXlsxFilePath, true);
            
            //------ レコード情報を書き込み ------

            using (var excel = new ExcelPackage(editXlsxFile))
            {
                var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var dimension = sheet.Dimension;

                // フィールド名取得.
                var fieldNames = ExcelUtility.GetRowValueTexts(sheet, fieldNameRow)
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
                    if (sheet.Cells.End.Row < recordRow)
                    {
                        sheet.InsertRow(recordRow, 1);
                    }

                    // セル情報コピー.
                    for (var column = 1; column < dimension.End.Column; column++)
                    {
                        CloneCellFormat(sheet, recordStartRow, recordRow, column);
                    }
                }
                 
                // レコード情報をセルに入力.
                for (var i = 0; i < records.Length; i++)
                {
                    var recordRow = recordStartRow + i;

                    var cellOption = cellOptions.FirstOrDefault(x => x.recordName == records[i].recordName);

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
                        var cell = sheet.Cells[recordRow, fieldColumn + 1];

                        // 値設定.
                        cell.Value = value;

                        // セルオプション情報追加.
                        if (cellOption != null && cellOption.cellInfos != null)
                        {
                            var cellInfo = cellOption.cellInfos.FirstOrDefault(x => x.fieldName.ToLower() == fieldName);
                            
                            SetCellInfos(cell, cellInfo);
                        }
                    }
                }

                // セルサイズを調整.

                var graphics = Graphics.FromImage(new Bitmap(1, 1));

                // 幅.
                for (var c = 1; c < dimension.End.Column; c++)
                {
                    var columnWidth = sheet.Column(c).Width;

                    for (var r = 1; r <= dimension.End.Row; r++)
                    {
                        var cell = sheet.Cells[r, c];

                        var width = CalcTextWidth(graphics, cell);

                        if (columnWidth < width)
                        {
                            columnWidth = width;
                        }
                    }

                    sheet.Column(c).Width = columnWidth;
                }

                // 高さ.
                for (var r = 1; r <= dimension.End.Row; r++)
                {
                    for (var c = 1; c <= dimension.End.Column; c++)
                    {
                        var cell = sheet.Cells[r, c];
                        
                        var height = CalcTextHeight(graphics, cell, (int)sheet.Column(c).Width);

                        if (sheet.Row(r).Height < height)
                        {
                            sheet.Row(r).Height = height;
                        }

                        cell.Style.WrapText = true;
                        cell.Style.ShrinkToFit = true;
                    }
                }

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

        private static void SetCellInfos(ExcelRange cell, CellOptionLoader.CellInfo cellInfo)
        {
            if (cellInfo == null) { return; }

            if (!string.IsNullOrEmpty(cellInfo.author) || !string.IsNullOrEmpty(cellInfo.comment))
            {
                cell.AddComment(cellInfo.comment, cellInfo.author);
            }

            if (!string.IsNullOrEmpty(cellInfo.fontColor))
            {
                cell.Style.Font.Color.SetColor(ColorTranslator.FromHtml(cellInfo.fontColor));
            }

            if (!string.IsNullOrEmpty(cellInfo.backgroundColor))
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(cellInfo.backgroundColor));
            }
        }

        private static double CalcTextWidth(Graphics graphics, ExcelRange cell)
        {
            if (string.IsNullOrEmpty(cell.Text)) { return 0.0; }

            var font = cell.Style.Font;

            var drawingFont = new Font(font.Name, font.Size);

            var size = graphics.MeasureString(cell.Text, drawingFont);

            return Convert.ToDouble(size.Width) / 5.7;
        }

        private static double CalcTextHeight(Graphics graphics, ExcelRange cell, int width)
        {
            if (string.IsNullOrEmpty(cell.Text)) { return 0.0; }

            var font = cell.Style.Font;

            var pixelWidth = Convert.ToInt32(width * 7.5);

            var drawingFont = new Font(font.Name, font.Size);

            var size = graphics.MeasureString(cell.Text, drawingFont, pixelWidth);

            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96 * 1.2, 409) + 2;
        }
    }
}
