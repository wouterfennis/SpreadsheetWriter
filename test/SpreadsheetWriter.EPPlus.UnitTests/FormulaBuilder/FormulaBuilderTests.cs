using System;
using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetWriter.Abstractions.Cell;
using SpreadsheetWriter.Abstractions.Formula;

namespace SpreadsheetWriter.EPPlus.UnitTests.Formula
{
    [TestClass]
    public class FormulaBuilderTests
    {
        private Fixture _fixture;
        private EPPlus.Formula.FormulaBuilder _formulaBuilder;

        [TestInitialize]
        public void Initialize()
        {
            _fixture = new Fixture();
            _formulaBuilder = new EPPlus.Formula.FormulaBuilder();
        }

        [TestMethod]
        public void AddCellAddress_WithoutAddress_ThrowsException()
        {
            // Arrange
            ICellAddress address = null;

            // Act
            Action action = () => _formulaBuilder.AddCellAddress(address);

            // Assert
            action.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void AddCellAddress_WithAddress_AddsAddressToFormula()
        {
            // Arrange
            CellAddress expectedAddress = CellAddress.Create($"A{_fixture.Create<int>()}");

            // Act
            IFormulaBuilder result = _formulaBuilder.AddCellAddress(expectedAddress);

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedAddress.ToString());
        }

        [TestMethod]
        public void AddDivisionSign_WithoutPreviousActions_AddsDivisionSignToFormula()
        {
            // Arrange
            var expectedFormula = "/";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddDivisionSign();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddMultiplicationSign_WithoutPreviousActions_AddsMultiplicationSignToFormula()
        {
            // Arrange
            var expectedFormula = "*";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddMultiplicationSign();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddSubtractionSign_WithoutPreviousActions_AddsSubtractionSignToFormula()
        {
            // Arrange
            var expectedFormula = "-";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddSubtractionSign();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddSummationSign_WithoutPreviousActions_AddsSummationSignToFormula()
        {
            // Arrange
            var expectedFormula = "+";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddSummationSign();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddOpenParenthesis_WithoutPreviousActions_AddsOpenParenthesisToFormula()
        {
            // Arrange
            var expectedFormula = "(";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddOpenParenthesis();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddClosingParenthesis_WithoutPreviousActions_AddsClosingParenthesisToFormula()
        {
            // Arrange
            var expectedFormula = ")";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddClosingParenthesis();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddEqualsSign_WithoutPreviousActions_AddsEqualsSignToFormula()
        {
            // Arrange
            var expectedFormula = "=";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddEqualsSign();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddConstantSign_WithoutPreviousActions_AddsConstantSignToFormula()
        {
            // Arrange
            var expectedFormula = "$";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddConstantSign();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddFormulaType_WithoutPreviousActions_AddsFormulaTypeToFormula()
        {
            // Arrange
            var expectedFormula = "COUNTIF";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddFormulaType(FormulaType.COUNTIF);

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddColon_WithoutPreviousActions_AddsColonToFormula()
        {
            // Arrange
            var expectedFormula = ":";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddColon();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddComma_WithoutPreviousActions_AddsColonToFormula()
        {
            // Arrange
            var expectedFormula = ",";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddComma();

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddCriteria_WithoutPreviousActions_AddsCriteriaToFormula()
        {
            // Arrange
            var expectedFormula = "\"<1\"";

            // Act
            IFormulaBuilder result = _formulaBuilder.AddCriteria("<1");

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void AddValue_WithoutPreviousActions_AddsValueToFormula()
        {
            // Arrange
            var expectedFormula = "1.09";

            // a
            IFormulaBuilder result = _formulaBuilder.AddValue(1.09);

            // Assert
            var formula = result.Build();
            formula.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void Build_WithoutPreviousActions_ReturnsEmptyString()
        {
            // Arrange
            var expectedFormula = string.Empty;

            // Act
            string result = _formulaBuilder.Build();

            // Assert
            result.Should().Be(expectedFormula);
        }

        [TestMethod]
        public void Build_WithoutPreviousActions_ReturnsFormulaString()
        {
            // Arrange
            var expectedFormula = "()";
            _formulaBuilder.AddOpenParenthesis();
            _formulaBuilder.AddClosingParenthesis();

            // Act
            string result = _formulaBuilder.Build();

            // Assert
            result.Should().Be(expectedFormula);
        }
    }
}
