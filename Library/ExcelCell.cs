
using System;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace Extensions
{
    public class ExcelCell
    {
        public sealed class Color
        {
            public string rgb = null;
            public eThemeSchemeColor? theme = null;
            public decimal? tint = null;
        }

        public string comment = null;

        public Color fontColor = null;

        public Color backgroundColor = null;

        public ExcelFillStyle patternType = ExcelFillStyle.Solid;
    }

    public static class ExcelCellUtility
    {
        public static bool ValidateAddress(ExcelWorksheet worksheet, int row, int column)
        {
            if (row <= 0 || column <= 0) { return false; }

            var endAddress = worksheet.Cells.End;

            if (endAddress.Row <= row || endAddress.Column <= column) { return false; }

            return true;
        }

        public static void Set<T>(ExcelWorksheet worksheet, int row, int column, T cellData) where T : ExcelCell
        {
            if (cellData == null) { return; }

            if (!ValidateAddress(worksheet, row, column)) { return; }

            var cell = worksheet.Cells[row, column];

            if (cell.Comment != null)
            {
                worksheet.Comments.Remove(cell.Comment);
            }

            if (!string.IsNullOrEmpty(cellData.comment))
            {
                cell.AddComment(cellData.comment, "REF");
            }

            cell.Style.Fill.PatternType = cellData.patternType;

            SetColor(cell.Style.Font.Color, cellData.fontColor);
            SetColor(cell.Style.Fill.BackgroundColor, cellData.backgroundColor);
        }

        public static T Get<T>(ExcelWorksheet worksheet, int row, int column) where T : ExcelCell
        {
            if (!ValidateAddress(worksheet, row, column)) { return null; }

            var cell = worksheet.Cells[row, column];

            var cellData = Activator.CreateInstance<T>();

            if (cell.Comment != null)
            {
                var comment = cell.Comment.Text;

                var author = cell.Comment.Author;

                var removeText = string.Format("{0}:", author);

                if (!string.IsNullOrEmpty(author) && comment.StartsWith(removeText))
                {
                    comment = comment.Substring(removeText.Length);
                }

                cellData.comment = comment.Trim('\n');
            }

            cellData.fontColor = GetFontColor(cell.Style.Font.Color);
            cellData.backgroundColor = GetBackgroundColor(cell.Style.Fill.BackgroundColor);
            cellData.patternType = cell.Style.Fill.PatternType;

            if (IsEmptyCellData(cellData)) { return null; }

            return cellData;
        }

        private static ExcelCell.Color GetFontColor(ExcelColor excelColor)
        {
            var color = new ExcelCell.Color()
            {
                rgb = string.IsNullOrEmpty(excelColor.Rgb) ? null : excelColor.Rgb,
                theme = excelColor.Theme,
            };

            if (color.theme.HasValue)
            {
                color.tint = excelColor.Tint;
            }

            if (IsEmptyColor(color)) { return null; }

            // 通常色の場合は無視.
            if (color.theme == eThemeSchemeColor.Text1 && color.tint == 0) { return null; }

            return color;
        }

        private static ExcelCell.Color GetBackgroundColor(ExcelColor excelColor)
        {
            var color = new ExcelCell.Color()
            {
                rgb = string.IsNullOrEmpty(excelColor.Rgb) ? null : excelColor.Rgb,
                theme = excelColor.Theme,
            };

            if (color.theme.HasValue)
            {
                color.tint = excelColor.Tint;
            }

            if (IsEmptyColor(color)) { return null; }

            return color;
        }

        private static void SetColor(ExcelColor excelColor, ExcelCell.Color color)
        {
            if (IsEmptyColor(color)) { return; }

            if (color.theme.HasValue)
            {
                excelColor.SetColor(color.theme.Value);

                if (color.tint.HasValue)
                {
                    excelColor.Tint = color.tint.Value;
                }
            }
            else
            {
                excelColor.SetColor(ColorTranslator.FromHtml("#" + color.rgb));
            }
        }

        private static bool IsEmptyCellData(ExcelCell cellData)
        {
            if (cellData == null) { return true; }

            var hasValue = false;

            hasValue |= !string.IsNullOrEmpty(cellData.comment);
            hasValue |= cellData.fontColor != null;
            hasValue |= cellData.backgroundColor != null;

            return !hasValue;
        }

        private static bool IsEmptyColor(ExcelCell.Color color)
        {
            if (color == null) { return true; }

            var hasValue = false;

            hasValue |= !string.IsNullOrEmpty(color.rgb);
            hasValue |= color.theme.HasValue;
            hasValue |= color.tint.HasValue;

            return !hasValue;
        }
    }
}
