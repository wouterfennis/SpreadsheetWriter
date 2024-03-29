﻿using System;
using System.IO;

namespace SpreadsheetWriter.Abstractions.File
{
    /// <summary>
    /// Result of saving files to the filesystem.
    /// </summary>
    public class SaveResult
    {
        /// <summary>
        /// Indicates if the save was a success.
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// The file location.
        /// </summary>
        public FileInfo FileInfo { get; set; }

        /// <summary>
        /// If there was an exception during the save. It is returned here.
        /// </summary>
        public Exception Exception { get; set; }
    }
}
