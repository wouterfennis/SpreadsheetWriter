using OfficeOpenXml;
using SpreadsheetWriter.Abstractions.Styling;
using System;
using System.Drawing;

namespace SpreadsheetWriter.EPPlus.Extensions
{
    /// <summary>
    ///  Extensions for the <see cref="ExcelRange"/>.
    /// </summary>
    internal static class ExcelRangeExtensions
    {
        /// <summary>
        /// Set the background color of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetBackgroundColor(this ExcelRange excelRange, Color color)
        {
            _ = excelRange?.Style?.Fill?.BackgroundColor ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            excelRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            excelRange.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// Set all the borders of the cell to the default values.
        /// </summary>
        /// <param name="excelRange"></param>
        public static void SetDefaultCellBorder(this ExcelRange excelRange)
        {
            excelRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, Color.LightGray);
        }

        /// <summary>
        /// Set the border of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetBorder(this ExcelRange excelRange, BorderDirection borderDirection, BorderStyle borderStyle, Color borderColor)
        {
            _ = excelRange?.Style?.Border ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            var mappedStyle = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
            switch (borderDirection)
            {
                case BorderDirection.Left:
                    excelRange.Style.Border.Left.Style = mappedStyle;
                    excelRange.Style.Border.Left.Color.SetColor(borderColor);
                    break;
                case BorderDirection.Right:
                    excelRange.Style.Border.Right.Style = mappedStyle;
                    excelRange.Style.Border.Right.Color.SetColor(borderColor);
                    break;
                case BorderDirection.Top:
                    excelRange.Style.Border.Top.Style = mappedStyle;
                    excelRange.Style.Border.Top.Color.SetColor(borderColor);
                    break;
                case BorderDirection.Bottom:
                    excelRange.Style.Border.Bottom.Style = mappedStyle;
                    excelRange.Style.Border.Bottom.Color.SetColor(borderColor);
                    break;
                case BorderDirection.Diagonal:
                    excelRange.Style.Border.Diagonal.Style = mappedStyle;
                    excelRange.Style.Border.Diagonal.Color.SetColor(borderColor);
                    break;
                case BorderDirection.DiagonalUp:
                    excelRange.Style.Border.Diagonal.Style = mappedStyle;
                    excelRange.Style.Border.Diagonal.Color.SetColor(borderColor);
                    excelRange.Style.Border.DiagonalUp = true;
                    break;
                case BorderDirection.DiagonalDown:
                    excelRange.Style.Border.Diagonal.Style = mappedStyle;
                    excelRange.Style.Border.Diagonal.Color.SetColor(borderColor);
                    excelRange.Style.Border.DiagonalDown = true;
                    break;
                case BorderDirection.Around:
                    excelRange.SetBorderAround(mappedStyle, borderColor);
                    break;
                case BorderDirection.None:
                    // No action.
                    break;
                default:
                    var exception = new InvalidOperationException(ExceptionMessages.UnknownBorderDirection);
                    exception.Data.Add(nameof(borderDirection), borderDirection);
                    throw exception;
            }
        }

        /// <summary>
        /// Shorthand for setting the border around.
        /// </summary>
        private static void SetBorderAround(this ExcelRange excelRange, OfficeOpenXml.Style.ExcelBorderStyle borderStyle, Color borderColor)
        {
            excelRange.Style.Border.BorderAround(borderStyle, borderColor);
        }

        /// <summary>
        /// Set the font color of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetFontColor(this ExcelRange excelRange, Color color)
        {
            _ = excelRange?.Style?.Font?.Color ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            excelRange.Style.Font.Color.SetColor(color);
        }

        /// <summary>
        /// Set the font size of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetFontSize(this ExcelRange excelRange, float size)
        {
            _ = excelRange?.Style?.Font ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            excelRange.Style.Font.Size = size;
        }

        /// <summary>
        /// Set the font bold of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetFontBold(this ExcelRange excelRange, bool isActive)
        {
            _ = excelRange?.Style?.Font ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            excelRange.Style.Font.Bold = isActive;
        }

        /// <summary>
        /// Set the horizontal alignment of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetHorizontalAlignment(this ExcelRange excelRange, HorizontalAlignment horizontalAlignment)
        {
            _ = excelRange?.Style ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            var mappedAlignment = (OfficeOpenXml.Style.ExcelHorizontalAlignment)horizontalAlignment;
            excelRange.Style.HorizontalAlignment = mappedAlignment;
        }

        /// <summary>
        /// Set the vertical alignment of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetVerticalAlignment(this ExcelRange excelRange, VerticalAlignment verticalAlignment)
        {
            _ = excelRange?.Style ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);
            var mappedAlignment = (OfficeOpenXml.Style.ExcelVerticalAlignment)verticalAlignment;
            excelRange.Style.VerticalAlignment = mappedAlignment;
        }

        /// <summary>
        /// Set the text rotation size of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetTextRotation(this ExcelRange excelRange, int rotation)
        {
            _ = excelRange?.Style ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            excelRange.Style.TextRotation = rotation;
        }

        /// <summary>
        /// Set the format of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetFormat(this ExcelRange excelRange, string format)
        {
            _ = excelRange?.Style?.Numberformat ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            excelRange.Style.Numberformat.Format = format;
        }
    }
}
