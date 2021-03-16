using System;
using System.Text.RegularExpressions;

namespace SpreadsheetWriter.Abstractions.Cell
{
    /// <summary>
    /// The definition of a cell address.
    /// </summary>
    public class CellAddress : ICellAddress
    {
        private const string CellAddressRegex = @"^([a-zA-Z]+)([1-9]+)$";

        /// <summary>
        /// The letter to state the column identifier.
        /// </summary>
        public ColumnLetter ColumnLetter { get; }

        /// <summary>
        /// The number to state the row identifier.
        /// </summary>
        public RowNumber RowNumber { get; }

        /// <inheritdoc/>
        public override string ToString()
        {
            return $"{ColumnLetter}{RowNumber}";
        }

        /// <summary>
        /// Constructor.
        /// </summary>
        private CellAddress(string columnLetter, string rowNumber)
        {
            ColumnLetter = ColumnLetter.Create(columnLetter);
            RowNumber = RowNumber.Create(rowNumber);
        }

        public static CellAddress Create(string cellAddress)
        {
            _ = cellAddress ?? throw new ArgumentNullException(nameof(cellAddress));
            Match match = Regex.Match(cellAddress, CellAddressRegex);

            if (match.Success)
            {
                string columnLetter = match.Groups[1].Value;
                string rowNumber = match.Groups[2].Value;
                return new CellAddress(columnLetter, rowNumber);
            }
            throw new InvalidOperationException("No valid input");
        }
    }
}
