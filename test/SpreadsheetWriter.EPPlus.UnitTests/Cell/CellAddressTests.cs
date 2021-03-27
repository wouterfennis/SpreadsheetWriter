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
        public void Create_WithValidInput_CreatesCellAddress()
        {
            // Arrange
            string input = "C1";

            // Act
            CellAddress result = CellAddress.Create(input);

            // Assert
            result.ColumnLetter.Value.Should().Be("C");
            result.RowNumber.Value.Should().Be("1");
        }

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
        public void Create_WithEmptyString_ThrowsException()
        {
            // Arrange 
            string input = "";

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void Create_WithOnlyColumn_ThrowsException()
        {
            // Arrange 
            string input = "C";

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void Create_WithOnlyRow_ThrowsException()
        {
            // Arrange 
            string input = "1";

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void Create_WithInvalidDivider_ThrowsException()
        {
            // Arrange 
            string input = "C/1";

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void Create_WithInvalidDivider2_ThrowsException()
        {
            // Arrange 
            string input = "C-1";

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }

        [TestMethod]
        public void Create_WithInvalidDivider3_ThrowsException()
        {
            // Arrange 
            string input = "C:1";

            // Act
            Action action = () => CellAddress.Create(input);

            // Assert
            action.Should().Throw<InvalidOperationException>();
        }
    }
}
