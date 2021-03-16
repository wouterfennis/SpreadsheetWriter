namespace SpreadsheetWriter.Abstractions.Cell
{
    public interface ICellAddress
    {
        ColumnLetter ColumnLetter { get; }
        RowNumber RowNumber { get; }

        string ToString();
    }
}