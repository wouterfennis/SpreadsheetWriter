using System;

namespace SpreadsheetWriter.Abstractions.File
{
    public class SaveResult
    {
        /// <summary>
        /// Indicates if the save went successful.
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public Exception Exception { get; set; }
    }
}
