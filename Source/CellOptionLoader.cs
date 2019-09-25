using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using YamlDotNet.Serialization;

namespace MasterConverter
{
    public static class CellOptionLoader
    {
        //----- params -----

        public class CellOption
        {
            public string recordName;
            public CellInfo[] cellInfos;
        }

        public class CellInfo
        {
            public string fieldName;
            public string author;
            public string comment;
            public string fontColor;
            public string backgroundColor;
        }

        //----- field -----

        //----- property -----

        //----- method -----

        /// <summary> セルオプション情報読み込み(.yaml) </summary>
        public static CellOption[] LoadYamlCellOptions(string yamlDirectory)
        {
            if (!Directory.Exists(yamlDirectory)) { return new CellOption[0]; }

            var optionFiles = Directory.EnumerateFiles(yamlDirectory, "*.*")
                .Where(x => Path.GetExtension(x) == Constants.CellOptionFileExtension);

            var list = new List<string>();

            foreach (var optionFile in optionFiles)
            {
                using (var file = new FileStream(optionFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var reader = new StreamReader(file, Encoding.UTF8))
                    {
                        list.Add(reader.ReadToEnd());
                    }
                }
            }

            var cellOptionList = new List<CellOption>();

            var deserializer = new DeserializerBuilder().Build();

            for (var i = 0; i < list.Count; i++)
            {
                var instance = deserializer.Deserialize<CellOption>(list[i]);

                cellOptionList.Add(instance);
            };

            return cellOptionList.ToArray();
        }

        /// <summary> セル情報読み込み(.xlsx) </summary>
        public static CellOption[] LoadXlsxRecordsCellOptions(string xlsxFilePath, RecordLoader.RecordData[] recordDatas)
        {
            var cellOptions = new List<CellOption>();

            if (!File.Exists(xlsxFilePath)) { return new CellOption[0]; }

            using (var excel = new ExcelPackage(new FileInfo(xlsxFilePath)))
            {
                var sheet = excel.Workbook.Worksheets.FirstOrDefault(x => x.Name == Constants.MasterSheetName);
                var address = sheet.Dimension;

                foreach (var recordData in recordDatas)
                {
                    var cellInfos = new List<CellInfo>();

                    for (var i = 0; i < recordData.values.Length; i++)
                    {
                        var recordValue = recordData.values[i];

                        if (recordValue == null) { continue; }

                        var r = recordData.index;
                        var c = address.Start.Column + i;

                        var cellInfo = GetCellInfo(sheet.Cells[r, c]);

                        if (cellInfo != null)
                        {
                            cellInfo.fieldName = recordValue.fieldName;
                            cellInfos.Add(cellInfo);
                        }
                    }

                    var option = new CellOption()
                    {
                        recordName = recordData.recordName,
                        cellInfos = cellInfos.ToArray(),
                    };

                    cellOptions.Add(option);
                }
            }

            return cellOptions.ToArray();
        }

        private static CellInfo GetCellInfo(ExcelRange cell)
        {
            CellInfo cellInfo = null;

            var author = cell.Comment != null ? cell.Comment.Author : null;
            var comment = cell.Comment != null ? cell.Comment.Text : null;

            var fontColor = GetColorCode(cell, cell.Style.Font.Color);
            var backgroundColor = GetColorCode(cell, cell.Style.Fill.BackgroundColor);

            var changed = false;

            changed |= !string.IsNullOrEmpty(author);
            changed |= !string.IsNullOrEmpty(comment);
            changed |= !string.IsNullOrEmpty(fontColor) && fontColor != "#FF000000";
            changed |= !string.IsNullOrEmpty(backgroundColor) && backgroundColor != "#FFFFFFFF";

            if (changed)
            {
                cellInfo = new CellInfo()
                {
                    author = author,
                    comment = comment,
                    fontColor = fontColor,
                    backgroundColor = backgroundColor,
                };
            }

            return cellInfo;
        }

        private static string GetColorCode(ExcelRange cell, ExcelColor color)
        {
            string colorCode = null;

            if (!string.IsNullOrEmpty(color.Rgb))
            {
                colorCode = "#" + color.Rgb;
            }

            if (!string.IsNullOrEmpty(color.Theme))
            {
                colorCode = null;

                throw new NotSupportedException(string.Format("Theme color not support.\n[{0}] {1}", cell.Address, cell.Text));
            }

            return colorCode;
        }        
    }
}
