using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.Abstractions;
using SpreadsheetWriter.EPPlus.Extensions;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;

namespace SpreadsheetWriter.EPPlus.UnitTests.ExcelSpreadsheetWriterTests
{
    [TestClass]
    public class WriteTests
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
        public void Write_WithValidValue_ReturnsInstanceOfSut()
        {
            // Arrange
            string value = _fixture.Create<string>();

            // Act
            ISpreadsheetWriter result = _sut.Write(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void Write_WithStringValue_InsertsStringInCurrentCell()
        {
            // Arrange
            string expectedValue = _fixture.Create<string>();

            // Act
            _sut.Write(expectedValue);

            // Assert
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Value.Should().Be(expectedValue);
        }

        [TestMethod]
        public void Write_WithDecimalValue_InsertsDecimalInCurrentCell()
        {
            // Arrange
            decimal expectedValue = _fixture.Create<decimal>();

            // Act
            _sut.Write(expectedValue);

            // Assert
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Value.Should().Be(expectedValue);
        }

        [TestMethod]
        public void Write_WithNull_ClearsExistingValueInCurrentCell()
        {
            // Arrange
            string expectedValue = null;
            _sut.Write("existing value");

            // Act
            _sut.Write(expectedValue);

            // Assert
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Value.Should().Be(expectedValue);
        }
    }
}
