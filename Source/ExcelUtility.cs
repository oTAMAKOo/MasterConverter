
using System;
using System.Linq;
using System.Collections.Generic;
using OfficeOpenXml;

namespace MasterConverter
{
    public static class ExcelUtility
    {
        /// <summary> 1行取得. </summary>
        public static IEnumerable<object> GetRowValues(ExcelWorksheet sheet, int row)
        {
            var address = sheet.Dimension;

            var values = new List<object>();

            for (var i = address.Start.Column; i <= address.End.Column; i++)
            {
                values.Add(sheet.Cells[row, i].Value);
            }

            return values;
        }

        /// <summary> 1行取得(文字列). </summary>
        public static IEnumerable<string> GetRowValueTexts(ExcelWorksheet sheet, int row)
        {
            var address = sheet.Dimension;

            var values = new List<string>();

            for (var i = address.Start.Column; i <= address.End.Column; i++)
            {
                values.Add(sheet.Cells[row, i].Text);
            }

            return values;
        }
    }
}
