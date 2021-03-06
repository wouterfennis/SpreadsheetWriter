using System;
using System.Drawing;
using OfficeOpenXml;
using SpreadsheetWriter.Abstractions.Styling;

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
        /// Set the border of an <see cref="ExcelRange"/>.
        /// </summary>
        public static void SetBorder(this ExcelRange excelRange, BorderDirection borderDirection, BorderStyle borderStyle)
        {
            _ = excelRange?.Style?.Border ?? throw new ArgumentException(ExceptionMessages.ExcelRangeNull);

            switch (borderDirection)
            {
                case BorderDirection.Left:
                    excelRange.Style.Border.Left.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
                    break;
                case BorderDirection.Right:
                    excelRange.Style.Border.Right.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;

                    break;
                case BorderDirection.Top:
                    excelRange.Style.Border.Top.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
                    break;
                case BorderDirection.Bottom:
                    excelRange.Style.Border.Bottom.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
                    break;
                case BorderDirection.Diagonal:
                    excelRange.Style.Border.Diagonal.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
                    break;
                case BorderDirection.DiagonalUp:
                    excelRange.Style.Border.Diagonal.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
                    excelRange.Style.Border.DiagonalUp = true;
                    break;
                case BorderDirection.DiagonalDown:
                    excelRange.Style.Border.Diagonal.Style = (OfficeOpenXml.Style.ExcelBorderStyle)borderStyle;
                    excelRange.Style.Border.DiagonalDown = true;
                    break;
                default:
                    var exception = new InvalidOperationException(ExceptionMessages.UnknownBorderDirection);
                    exception.Data.Add(nameof(borderDirection), borderDirection);
                    throw exception;
            }
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
