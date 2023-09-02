namespace SpreadsheetWriter.Abstractions.Formula
{
    /// <summary>
    /// The types of formulas supported.
    /// </summary>
    public enum FormulaType
    {
        /// <summary>
        /// Calculates the sum of the selected cells.
        /// </summary>
        SUM,

        /// <summary>
        /// Calculates the average of the selected cells.
        /// </summary>
        AVERAGE,

        /// <summary>
        /// Calculates occurences of value
        /// </summary>
        COUNT,

        /// <summary>
        /// Calculates occurences of value, if it meets the condition.
        /// </summary>
        COUNTIF
    }
}