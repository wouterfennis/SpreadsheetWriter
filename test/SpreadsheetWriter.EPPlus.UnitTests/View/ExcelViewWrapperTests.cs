using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;
using SpreadsheetWriter.EPPlus.View;
using System.Linq;

namespace SpreadsheetWriter.EPPlus.UnitTests.View
{
    [TestClass]
    public class ExcelViewWrapperTests
    {
        private ExcelViewWrapper _sut;
        private ExcelWorksheet _worksheet;

        [TestInitialize]
        public void Initialize()
        {
            _worksheet = ExcelTestBuilder.CreateExcelWorksheet();
            _sut = new ExcelViewWrapper(_worksheet.View);
        }

        [TestMethod]
        public void FreezePanes_WithOneCellSelection_FreezesCellSelection()
        {
            // Arrange

            // Act
            _sut.FreezePanes(1, 1);

            // Assert
            _worksheet.View.Panes.Length.Should().Be(1);
            _worksheet.View.Panes.First().SelectedRange.Should().Be("A1");  
        }

        [TestMethod]
        public void FreezePanes_WithMultipleCellsSelection_FreezesCellSelection()
        {
            // Arrange

            // Act
            _sut.FreezePanes(10, 2);

            // Assert
            _worksheet.View.Panes.Length.Should().Be(3);
            _worksheet.View.Panes[0].SelectedRange.Should().Be("B1");
            _worksheet.View.Panes[1].SelectedRange.Should().Be("A10");
            _worksheet.View.Panes[2].SelectedRange.Should().Be("A1");
        }
    }
}
