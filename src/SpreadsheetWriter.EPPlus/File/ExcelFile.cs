using System;
using System.IO;
using System.Threading.Tasks;
using OfficeOpenXml;
using SpreadsheetWriter.Abstractions;
using SpreadsheetWriter.Abstractions.File;

namespace SpreadsheetWriter.EPPlus.File
{
    /// <inheritdoc/>
    public sealed class ExcelFile : ISpreadsheetFile
    {
        private ExcelPackage _excelPackage;
        private ISpreadsheetWriter _writer;

        public ExcelFile(string fileLocation, Metadata metadata)
        {
            var fileInfo = new FileInfo(fileLocation);
            _excelPackage = new ExcelPackage(fileInfo);
            _excelPackage.Workbook.Properties.Author = metadata.Author;
            _excelPackage.Workbook.Properties.Title = metadata.Title;
            _excelPackage.Workbook.Properties.Subject = metadata.Subject;
            _excelPackage.Workbook.Properties.Created = metadata.Created;

            ExcelWorksheet worksheet = _excelPackage.Workbook.Worksheets.Add(metadata.Title);
            _writer = new ExcelSpreadsheetWriter(worksheet);
        }

        /// <inheritdoc/>
        public ISpreadsheetWriter GetSpreadsheetWriter()
        {
            return _writer;
        }

        /// <inheritdoc/>
        public async Task<SaveResult> SaveAsync()
        {
            Exception exceptionDuringSave = null;
            bool isSuccess = false;
            try
            {
                await _excelPackage.SaveAsync();
                isSuccess = true;
            }
            catch (Exception exception)
            {
                exceptionDuringSave = exception;
            }

            return new SaveResult
            {
                IsSuccess = isSuccess,
                Exception = exceptionDuringSave
            };
        }

        /// <inheritdoc/>
        public void Dispose()
        {
            _excelPackage.Dispose();
        }
    }
}
