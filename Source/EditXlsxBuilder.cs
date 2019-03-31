
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public static void Build(string originXlsxFilePath, SerializeClass serializeClass, RecordLoader.RecordData[] records, int fieldNameRow, int recordStartRow)
        {
            var editXlsxFilePath = PathUtility.Combine(Path.GetDirectoryName(originXlsxFilePath), Constants.MasterFileName);

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
                for (var i = recordStartRow; i < records.Length; i++)
                {
                    foreach (var recordValue in records[i].values)
                    {
                        var fieldName = recordValue.fieldName.ToLower();
                        var column = fieldIndexDictionary.GetValueOrDefault(fieldName, -1);

                        if (column == -1) { continue; }

                        if (sheet.Cells.End.Row < i)
                        {
                            sheet.InsertRow(i, 1);
                        }

                        // Excelのセルは1開始なので1加算.
                        sheet.Cells[i, column + 1].Value = recordValue.value;
                    }
                }

                // セルサイズを調整.
                sheet.Cells.AutoFitColumns();

                // 保存.
                excel.Save();
            }
        }
    }
}
