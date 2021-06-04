using System;
using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;

namespace SpreadsheetWriter.EPPlus.UnitTests.ExcelSpreadsheetWriterTests
{
    [TestClass]
    public class ConstructorTests
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
        public void Constructor_WithoutWorksheet_ThrowsException()
        {
            // Arrange
            ExcelWorksheet worksheet = null;

            // Act
            Action action = () => new ExcelSpreadsheetWriter(worksheet);

            // Assert
            action.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void Constructor_WithValidWorksheet_SetsCurrentCellToTopLeft()
        {
            // Arrange
            var worksheet = ExcelTestBuilder.CreateExcelWorksheet();

            // Act
            _sut = new ExcelSpreadsheetWriter(worksheet);

            // Assert
            _sut.CurrentPosition.X.Should().Be(1);
            _sut.CurrentPosition.Y.Should().Be(1);
        }
    }
}