using System;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SpreadsheetWriter.Abstractions.Cell;

namespace SpreadsheetWriter.EPPlus.UnitTests.Cell
{
    [TestClass]
    public class CellAddressTests
    {
        [TestMethod]
        public void Create_WithoutInput_ThrowsException()
        {
            // Arrange
            string input = null;

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<ArgumentNullException>();
        }

        [TestMethod]
        //[DataRow("")]
        //[DataRow("C")]
        //[DataRow("1")]
        //[DataRow("C/1")]
        //[DataRow("C:1")]
        //[DataRow("C-1")]
        public void Create_WithInvalidInput_ThrowsException(string input)
        {
            // Arrange in datarow

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }
    }
}
