
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using Extensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MasterConverter
{
    public class EditXlsxBuilder
    {
        //----- params -----

        //----- field -----

        //----- property -----

        //----- method -----

        public static void Build(string originXlsxFilePath, SerializeClass serializeClass, RecordLoader.RecordData[] records, int fieldNameRow, int recordStartRow)
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

                // フィールド名のセル位置を辞書化.
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

                // レコード情報をセルに入力.
                for (var i = 0; i < records.Length; i++)
                {
                    var recordRow = recordStartRow + i;

                    foreach (var recordValue in records[i].values)
                    {
                        var fieldName = recordValue.fieldName.ToLower();
                        var column = fieldIndexDictionary.GetValueOrDefault(fieldName, -1);
                        
                        if (column == -1) { continue; }

                        if (sheet.Cells.End.Row < recordRow)
                        {
                            sheet.InsertRow(recordRow, 1);
                        }

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
                        var cell = sheet.Cells[recordRow, column + 1];
                        
                        cell.Value = value;
                    }                    
                }

                // セルサイズを調整.

                sheet.Cells[dimension.Address].AutoFitColumns();

                for (var c = 1; c < dimension.End.Column; c++)
                {
                    sheet.Column(c).Width *= 1.5f;
                }

                for (var r = 1; r < dimension.End.Row; r++)
                {
                    for (var c = 1; c < dimension.End.Column; c++)
                    {
                        var cell = sheet.Cells[r, c];
                        
                        var height = MeasureTextHeight(cell.Text, cell.Style.Font, (int)sheet.Column(c).Width);

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

        public static double MeasureTextHeight(string text, ExcelFont font, int width)
        {
            if (string.IsNullOrEmpty(text)) return 0.0;
            var bitmap = new Bitmap(1, 1);
            var graphics = Graphics.FromImage(bitmap);

            var pixelWidth = Convert.ToInt32(width * 7.5); //7.5 pixels per excel column width
            var drawingFont = new Font(font.Name, font.Size);
            var size = graphics.MeasureString(text, drawingFont, pixelWidth);

            //72 DPI and 96 points per inch.  Excel height in points with max of 409 per Excel requirements.
            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96, 409);
        }
    }
}
