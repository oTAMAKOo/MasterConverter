
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

                        var type = recordValue.value.GetType();

                        if (type.IsArray)
                        {
                            var array = recordValue.value as IEnumerable;

                            var valueTexts = new List<string>();

                            foreach (var item in array)
                            {
                                valueTexts.Add(item.ToString());
                            }

                            value = string.Format("[{0}]", string.Join(",", valueTexts));
                        }
                        else
                        {
                            value = recordValue.value;
                        }

                        // Excelのセルは1開始なので1加算.
                        sheet.Cells[recordRow, column + 1].Value = value;
                    }                    
                }

                // セルサイズを調整.
                sheet.Cells[dimension.Address].AutoFitColumns();

                for (var i = 1; i < dimension.End.Column; i++)
                {
                    sheet.Column(i).Width *= 1.5f;
                }

                // 保存.
                excel.Save();
            }
        }
    }
}
