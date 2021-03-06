using System.Drawing;
using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using OfficeOpenXml;
using SpreadsheetWriter.Abstractions;
using SpreadsheetWriter.Abstractions.Formula;
using SpreadsheetWriter.EPPlus.Extensions;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;
using SpreadsheetWriter.EPPlus.UnitTests.Utilities;

namespace SpreadsheetWriter.EPPlus.UnitTests.ExcelSpreadsheetWriterTests
{
    [TestClass]
    public class FormulaTests
    {
        private ExcelSpreadsheetWriter _sut;
        private Fixture _fixture;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _fixture = new Fixture();
            _worksheet = ExcelTestBuilder.CreateExcelWorksheet();
            _sut = new ExcelSpreadsheetWriter(_worksheet);
        }

        [TestMethod]
        public void PlaceStandardFormula_WithValidValue_ReturnsInstanceOfSut()
        {
            // Arrange
            var startPosition = new Point(_fixture.Create<short>(), _fixture.Create<short>());
            var endPosition = new Point(_fixture.Create<short>(), _fixture.Create<short>());
            var formulaType = _fixture.Create<FormulaType>();

            // Act
            ISpreadsheetWriter result = _sut.PlaceStandardFormula(startPosition, endPosition, formulaType);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void PlaceStandardFormula_WithValidPoints_PlacesFormulaInCurrentCell()
        {
            // Arrange
            var startPosition = new Point(_fixture.Create<short>(), _fixture.Create<short>());
            var endPosition = new Point(_fixture.Create<short>(), _fixture.Create<short>());
            var formulaType = _fixture.Create<FormulaType>();

            // Act
            _sut.PlaceStandardFormula(startPosition, endPosition, formulaType);

            // Assert
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            var startAddress = $"{ExcelColumnUtility.GetExcelColumnName(startPosition.X)}{startPosition.Y}";
            var endAddress = $"{ExcelColumnUtility.GetExcelColumnName(endPosition.X)}{endPosition.Y}";

            cell.Formula.Should().Contain($"={formulaType}({startAddress}:{endAddress})");
        }

        [TestMethod]
        public void PlaceCustomFormula_WithValidValue_ReturnsInstanceOfSut()
        {
            // Arrange
            var expectedFormula = _fixture.Create<string>();
            var formulaBuilder = new Mock<IFormulaBuilder>();
            formulaBuilder.Setup(x => x.Build())
                .Returns(expectedFormula);

            // Act
            ISpreadsheetWriter result = _sut.PlaceCustomFormula(formulaBuilder.Object);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void PlaceCustomFormula_WithFormulaBuilder_PlacesFormulaToCurrentCell()
        {
            // Arrange
            var expectedFormula = _fixture.Create<string>();
            var formulaBuilder = new Mock<IFormulaBuilder>();
            formulaBuilder.Setup(x => x.Build())
                .Returns(expectedFormula);

            // Act
            ISpreadsheetWriter result = _sut.PlaceCustomFormula(formulaBuilder.Object);

            // Assert
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Formula.Should().Be(expectedFormula);
        }
    }
}
