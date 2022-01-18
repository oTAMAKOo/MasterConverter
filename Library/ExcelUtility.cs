
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using OfficeOpenXml;

namespace Extensions
{
    public static class ExcelUtility
    {
        private static Graphics graphics = null;

        private static List<Font> fonts = new List<Font>();

        public static T ConvertValue<T>(object value)
        {
            if (value is T) { return (T)value; }

            try
            {
                return (T)Convert.ChangeType(value, typeof(T));
            }
            catch (InvalidCastException)
            {
                return default(T);
            }
        }


        public static T ConvertValue<T>(object[] values, int index)
        {
            if (index < 0 || values.Length <= index)
            {
                throw new ArgumentOutOfRangeException();
            }

            var value = values[index];

            return ConvertValue<T>(value);
        }

        /// <summary> 1行取得. </summary>
        public static IEnumerable<object> GetRowValues(ExcelWorksheet sheet, int row)
        {
            var address = sheet.Dimension;

            var values = new List<object>();

            for (var i = address.Start.Column; i <= address.End.Column; i++)
            {
                var value = sheet.GetValue(row, i);

                values.Add(value);
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

        public static void FitColumnSize(ExcelWorksheet worksheet, ExcelRange range, int? minSize = null, int? maxSize = null,
                                         Func<int, int, string, bool> wrapTextCallback = null, 
                                         Func<int, int, string, bool> shrinkToFitCallback = null)
        {
            if (graphics == null)
            {
                graphics = Graphics.FromImage(new Bitmap(1, 1));
            }

            var min = minSize.HasValue ? minSize.Value : double.MinValue;
            var max = maxSize.HasValue ? maxSize.Value : double.MaxValue;

            for (var c = range.Start.Column; c < range.End.Column; c++)
            {
                worksheet.Column(c).AutoFit();

                var columnWidth = worksheet.Column(c).Width;

                for (var r = range.Start.Row; r <= range.End.Row; r++)
                {
                    var cell = worksheet.Cells[r, c];

                    if (string.IsNullOrEmpty(cell.Text)) { continue; }

                    cell.Style.WrapText = wrapTextCallback != null && wrapTextCallback.Invoke(r, c, cell.Text);
                    cell.Style.ShrinkToFit = shrinkToFitCallback != null && shrinkToFitCallback.Invoke(r, c, cell.Text);

                    var width = CalcTextWidth(graphics, cell);

                    if (columnWidth < width)
                    {
                        columnWidth = width;
                    }
                }

                worksheet.Column(c).Width = Math.Min(Math.Max(columnWidth, min), max);
            }
        }

        public static void FitRowSize(ExcelWorksheet worksheet, ExcelRange range, int? minSize = null, int? maxSize = null)
        {
            if (graphics == null)
            {
                graphics = Graphics.FromImage(new Bitmap(1, 1));
            }

            var min = minSize.HasValue ? minSize.Value : double.MinValue;
            var max = maxSize.HasValue ? maxSize.Value : double.MaxValue;

            for (var r = range.Start.Row; r <= range.End.Row; r++)
            {
                for (var c = range.Start.Column; c <= range.End.Column; c++)
                {
                    var cell = worksheet.Cells[r, c];

                    if (string.IsNullOrEmpty(cell.Text)) { continue; }

                    var columnWidth = (int)worksheet.Column(c).Width;

                    var height = CalcTextHeight(graphics, cell, columnWidth);

                    if (worksheet.Row(r).Height < height)
                    {
                        worksheet.Row(r).Height = Math.Min(Math.Max(height, min), max);
                    }
                }
            }
        }

        private static double CalcTextWidth(Graphics graphics, ExcelRange cell)
        {
            if (string.IsNullOrEmpty(cell.Text)) { return 0.0; }

            var font = cell.Style.Font;

            Font drawingFont;

            lock (fonts)
            {
                drawingFont = fonts.FirstOrDefault(x => x.Name == font.Name && x.Size == font.Size);

                if (drawingFont == null)
                {
                    drawingFont = new Font(font.Name, font.Size);

                    fonts.Add(drawingFont);
                }
            }

            SizeF size;

            lock (graphics)
            {
                size = graphics.MeasureString(cell.Text, drawingFont);
            }

            return Convert.ToDouble(size.Width) / 5.7;
        }

        private static double CalcTextHeight(Graphics graphics, ExcelRange cell, int width)
        {
            if (string.IsNullOrEmpty(cell.Text)) { return 0.0; }

            var font = cell.Style.Font;

            Font drawingFont;
            
            lock (fonts)
            {
                drawingFont = fonts.FirstOrDefault(x => x.Name == font.Name && x.Size == font.Size);

                if (drawingFont == null)
                {
                    drawingFont = new Font(font.Name, font.Size);

                    fonts.Add(drawingFont);
                }
            }

            var pixelWidth = Convert.ToInt32(width * 7.5);

            SizeF size;

            lock (graphics)
            {
                size = graphics.MeasureString(cell.Text, drawingFont, pixelWidth);
            }

            return Math.Min(Convert.ToDouble(size.Height) * 72 / 96 * 1.2, 409) + 2;
        }
    }
}
