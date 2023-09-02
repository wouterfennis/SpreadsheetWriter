using System;
using System.Globalization;
using System.Text;
using SpreadsheetWriter.Abstractions.Cell;
using SpreadsheetWriter.Abstractions.Formula;

namespace SpreadsheetWriter.EPPlus.Formula
{
    /// <inheritdoc/>
    public class FormulaBuilder : IFormulaBuilder
    {
        private readonly StringBuilder _stringBuilder;

        public FormulaBuilder()
        {
            _stringBuilder = new StringBuilder();
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddCellAddress(ICellAddress cellAddress)
        {
            _ = cellAddress ?? throw new ArgumentNullException(nameof(cellAddress));

            _stringBuilder.Append(cellAddress);
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddCellColumnLetter(ColumnLetter columnLetter)
        {
            _ = columnLetter ?? throw new ArgumentNullException(nameof(columnLetter));

            _stringBuilder.Append(columnLetter);
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddRowNumber(RowNumber rowNumber)
        {
            _ = rowNumber ?? throw new ArgumentNullException(nameof(rowNumber));

            _stringBuilder.Append(rowNumber);
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddDivisionSign()
        {
            _stringBuilder.Append("/");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddMultiplicationSign()
        {
            _stringBuilder.Append("*");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddSubtractionSign()
        {
            _stringBuilder.Append("-");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddSummationSign()
        {
            _stringBuilder.Append("+");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddOpenParenthesis()
        {
            _stringBuilder.Append("(");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddClosingParenthesis()
        {
            _stringBuilder.Append(")");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddEqualsSign()
        {
            _stringBuilder.Append("=");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddConstantSign()
        {
            _stringBuilder.Append("$");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddFormulaType(FormulaType formulaType)
        {
            _stringBuilder.Append(formulaType.ToString());
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddColon()
        {
            _stringBuilder.Append(":");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddComma()
        {
            _stringBuilder.Append(",");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddCriteria(string criteria)
        {
            _stringBuilder.Append($"\"{criteria}\"");
            return this;
        }

        /// <inheritdoc/>
        public IFormulaBuilder AddValue(double value)
        {
            _stringBuilder.Append(value.ToString(CultureInfo.InvariantCulture));
            return this;
        }

        /// <inheritdoc/>
        public string Build()
        {
            return _stringBuilder.ToString();
        }
    }
}
