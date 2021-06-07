using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using SpreadsheetWriter.Abstractions;
using SpreadsheetWriter.Abstractions.Cell;
using SpreadsheetWriter.Abstractions.Formula;
using SpreadsheetWriter.EPPlus.Extensions;
using System;
using System.Drawing;
using System.Globalization;

namespace SpreadsheetWriter.EPPlus
{
    /// <summary>
    /// Spreadsheet writer for the OfficeOpenXml library.
    /// </summary>
    public class ExcelSpreadsheetWriter : SpreadsheetWriterBase
    {
        private const int DefaultXPosition = 1;
        private const int DefaultYPosition = 1;
        private readonly ExcelWorksheet _excelWorksheet;

        private ExcelRange CurrentCell { get => _excelWorksheet.GetCell(CurrentPosition); }

        public ExcelSpreadsheetWriter(ExcelWorksheet excelWorksheet) : base(DefaultXPosition, DefaultYPosition)
        {
            _excelWorksheet = excelWorksheet ?? throw new ArgumentNullException(nameof(excelWorksheet));
            CurrentPosition = new Point(DefaultXPosition, DefaultYPosition);
        }

        /// <inheritdoc/>
        public override ICellRange GetCellRange(Point position)
        {
            return new ExcelRangeWrapper(_excelWorksheet.GetCell(position));
        }

        /// <inheritdoc/>
        public override ISpreadsheetWriter Write(decimal value)
        {
            WriteInternal(value);
            return this;
        }

        /// <inheritdoc/>
        public override ISpreadsheetWriter Write(string value)
        {
            WriteInternal(value);
            return this;
        }

        private void WriteInternal(object value)
        {
            ApplyStyling();
            CurrentCell.Value = value;
        }

        /// <inheritdoc/>
        public override ISpreadsheetWriter ApplyStyling()
        {
            CurrentCell.SetBackgroundColor(CurrentBackgroundColor);
            CurrentCell.SetFontColor(CurrentFontColor);
            CurrentCell.SetTextRotation(CurrentTextRotation);
            CurrentCell.SetFontSize(CurrentFontSize);
            CurrentCell.SetFontBold(IsCurrentFontBold);
            CurrentCell.SetFormat(CurrentFormat);
            CurrentCell.SetHorizontalAlignment(CurrentHorizontalAlignment);
            CurrentCell.SetVerticalAlignment(CurrentVerticalAlignment);
            CurrentCell.SetDefaultCellBorder();
            CurrentCell.SetBorder(CurrentBorderDirection, CurrentBorderStyle, CurrentBorderColor);

            return this;
        }

        /// <summary>
        /// Set the vertical alignment of an <see cref="ExcelRange"/>.
        /// </summary>
        public override ISpreadsheetWriter PlaceLessThanRule(double lessThanValue, Color fillColor)
        {
            IExcelConditionalFormattingLessThan rule = CurrentCell.ConditionalFormatting.AddLessThan();
            rule.Formula = lessThanValue.ToString(CultureInfo.InvariantCulture);
            rule.Style.Fill.BackgroundColor.Color = fillColor;
            return this;
        }

        /// <inheritdoc/>
        public override ISpreadsheetWriter PlaceStandardFormula(Point startPosition, Point endPosition, FormulaType formulaType)
        {
            ApplyStyling();

            var startCell = _excelWorksheet.GetCell(startPosition);
            var endCell = _excelWorksheet.GetCell(endPosition);
            var resultCell = _excelWorksheet.GetCell(CurrentPosition);

            var formula = $"={formulaType}({startCell.Address}:{endCell.Address})";
            resultCell.Formula = formula;

            return this;
        }

        /// <inheritdoc/>
        public override ISpreadsheetWriter PlaceCustomFormula(IFormulaBuilder formulaBuilder)
        {
            ApplyStyling();

            var resultCell = _excelWorksheet.GetCell(CurrentPosition);

            resultCell.Formula = formulaBuilder.Build();
            return this;
        }
    }
}
