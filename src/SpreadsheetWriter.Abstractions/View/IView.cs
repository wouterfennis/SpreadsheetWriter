namespace SpreadsheetWriter.Abstractions.View
{
    /// <summary>
    /// Abstraction for a view.
    /// </summary>
    public interface IView
    {
        /// <summary>
        /// Freeze the columns/rows to left and above the cell
        /// </summary>
        void FreezePanes(int row, int column);
    }
}
