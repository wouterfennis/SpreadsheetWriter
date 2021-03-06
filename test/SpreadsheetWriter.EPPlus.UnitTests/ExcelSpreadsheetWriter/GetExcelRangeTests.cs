using System.Drawing;
using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;

namespace SpreadsheetWriter.EPPlus.UnitTests.ExcelSpreadsheetWriterTests
{
    [TestClass]
    public class GetExcelRangeTests
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
        public void GetCellRange_WithValidPoint_ReturnsMatchingExcelRange()
        {
            // Arrange
            var point = new Point(_fixture.Create<short>(), _fixture.Create<short>());

            // Act
            var result = _sut.GetCellRange(point);

            // Assert
            result.Should().NotBeNull();
            result.Address.Should().Contain(point.Y.ToString());
        }
    }
}