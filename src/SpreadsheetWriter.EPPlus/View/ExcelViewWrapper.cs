using OfficeOpenXml;
using SpreadsheetWriter.Abstractions.View;
using System;

namespace SpreadsheetWriter.EPPlus.View
{
    /// <summary>
    /// Wrapper for the ExcelWorksheetView class.
    /// </summary>
    public class ExcelViewWrapper : IView
    {
        private readonly ExcelWorksheetView _excelView;

        public ExcelViewWrapper(ExcelWorksheetView excelView)
        {
            _excelView = excelView ?? throw new ArgumentNullException(nameof(excelView));
        }

        /// <inheritdoc/>
        public void FreezePanes(int row, int column)
        {
            _excelView.FreezePanes(row, column);
        }
    }
}
