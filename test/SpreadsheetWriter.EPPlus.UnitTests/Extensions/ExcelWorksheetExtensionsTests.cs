using System;
using System.Drawing;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.EPPlus.Extensions;

namespace SpreadsheetWriter.EPPlus.UnitTests.Extensions
{
    [TestClass]
    public class ExcelWorksheetExtensionsTests
    {
        [TestMethod]
        public void GetCell_WithExcelRangeNull_ThrowsException()
        {
            // Arrange
            ExcelWorksheet excelWorksheet = null;

            // Act
            Action action = () => excelWorksheet.GetCell(new Point());

            // Assert
            action.Should().Throw<ArgumentException>();
        }
    }
}
