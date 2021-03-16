using System;
using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.Abstractions.Cell;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;

namespace SpreadsheetWriter.EPPlus.UnitTests
{
    [TestClass]
    public class ExcelRangeWrapperTests
    {
        private Fixture _fixture;

        [TestInitialize]
        public void Initialize()
        {
            _fixture = new Fixture();
        }

        [TestMethod]
        public void Constructor_WithoutExcelRange_ThrowsException()
        {
            // Arrange
            ExcelRange excelRange = null;

            // Act
            Action action = () => new ExcelRangeWrapper(excelRange);

            // Assert
            action.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        public void Address_WithValidExcelRange_ReturnsAddressOfExcelRange()
        {
            // Arrange
            var excelRange = ExcelTestBuilder.CreateExcelRange();
            var sut = new ExcelRangeWrapper(excelRange);

            // Act
            ICellAddress result = sut.Address;

            // Assert
            result.ToString().Should().Be(excelRange.Address);
        }

        [TestMethod]
        public void Value_WithValidExcelRange_ReturnsValueOfExcelRange()
        {
            // Arrange
            var expectedValue = _fixture.Create<string>();
            var excelRange = ExcelTestBuilder.CreateExcelRange();
            excelRange.Value = expectedValue;
            var sut = new ExcelRangeWrapper(excelRange);

            // Act
            string result = sut.Value;

            // Assert
            result.Should().Be((string)excelRange.Value);
        }
    }
}
