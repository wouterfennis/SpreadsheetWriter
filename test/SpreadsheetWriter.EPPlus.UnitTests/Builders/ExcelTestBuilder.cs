using OfficeOpenXml;

namespace SpreadsheetWriter.EPPlus.UnitTests.Builders
{
    /// <summary>
    /// Builder to create test instances of Excel spreadsheets.
    /// </summary>
    public static class ExcelTestBuilder
    {
        /// <summary>
        /// Returns a default <see cref="ExcelRange"/>.
        /// </summary>
        public static ExcelRange CreateExcelRange()
        {
            var worksheet = CreateExcelWorksheet();
            return worksheet.Cells[1, 1];
        }

        /// <summary>
        /// Returns a default <see cref="ExcelWorksheet"/>.
        /// </summary>
        public static ExcelWorksheet CreateExcelWorksheet()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var excelPackage = new ExcelPackage();
            return excelPackage.Workbook.Worksheets.Add("test");
        }
    }
}
