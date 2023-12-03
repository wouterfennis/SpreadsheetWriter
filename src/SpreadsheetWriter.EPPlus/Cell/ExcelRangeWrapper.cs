using System;
using OfficeOpenXml;
using SpreadsheetWriter.Abstractions.Cell;

namespace SpreadsheetWriter.EPPlus.Cell
{
    /// <inheritdoc/>
    public class ExcelRangeWrapper : ICellRange
    {
        private readonly ExcelRange _excelRange;

        public ExcelRangeWrapper(ExcelRange excelRange)
        {
            _excelRange = excelRange ?? throw new ArgumentNullException(nameof(excelRange));
        }

        /// <inheritdoc/>
        public ICellAddress Address => CellAddress.Create(_excelRange.Address);

        /// <inheritdoc/>
        public string Value => (string)_excelRange.Value;
    }
}
