using System;
using System.Drawing;
using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using SpreadsheetWriter.Abstractions.Styling;
using SpreadsheetWriter.EPPlus.Extensions;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;

namespace SpreadsheetWriter.EPPlus.UnitTests.Extensions
{
    [TestClass]
    public class ExcelRangeExtensionsTests
    {
        private Fixture _fixture;

        [TestInitialize]
        public void Initialize()
        {
            _fixture = new Fixture();
        }

        [TestMethod]
        public void SetBackgroundColor_WithExcelRangeNull_ThrowsException()
        {
            // Arrange
            ExcelRange excelRange = null;

            // Act
            Action action = () => excelRange.SetBackgroundColor(_fixture.Create<Color>());

            // Assert
            action.Should().Throw<ArgumentException>();
        }

        [TestMethod]
        public void SetBorder_WithExcelRangeNull_ThrowsException()
        {
            // Arrange
            ExcelRange excelRange = null;

            // Act
            Action action = () => excelRange.SetBorder(_fixture.Create<BorderDirection>(), _fixture.Create<BorderStyle>());

            // Assert
            action.Should().Throw<ArgumentException>();
        }

        [TestMethod]
        public void SetBorder_WithBorderTop_SetStyleOnTopBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.Top, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Top.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
        }

        [TestMethod]
        public void SetBorder_WithBorderLeft_SetStyleOnLeftBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.Left, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Left.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
        }

        [TestMethod]
        public void SetBorder_WithBorderRight_SetStyleOnRightBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.Right, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Right.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
        }

        [TestMethod]
        public void SetBorder_WithBottom_SetStyleOnBottomBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.Bottom, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Bottom.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
        }

        [TestMethod]
        public void SetBorder_WithDiagonal_SetStyleOnDiagonalBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.Diagonal, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Diagonal.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
        }

        [TestMethod]
        public void SetBorder_WithDiagonalDown_SetStyleOnDiagnoalDownBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.DiagonalDown, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Diagonal.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
            excelRange.Style.Border.DiagonalDown.Should().BeTrue();
        }

        [TestMethod]
        public void SetBorder_WithDiagonalUp_SetStyleOnDiagonalUpBorder()
        {
            // Arrange
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();
            BorderStyle expectedBorderStyle = _fixture.Create<BorderStyle>();

            // Act
            excelRange.SetBorder(BorderDirection.DiagonalUp, expectedBorderStyle);

            // Assert
            excelRange.Style.Border.Diagonal.Style.Should().Be((ExcelBorderStyle)expectedBorderStyle);
            excelRange.Style.Border.DiagonalUp.Should().BeTrue();
        }

        [TestMethod]
        public void SetFontBold_WithActive_SetsBoldStyle()
        {
            // Arrange
            bool expectedToBeBold = true;
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();

            // Act
            excelRange.SetFontBold(expectedToBeBold);

            // Assert
            excelRange.Style.Font.Bold.Should().BeTrue();
        }

        [TestMethod]
        public void SetFontBold_WithInactive_SetsBoldStyle()
        {
            // Arrange
            bool expectedToBeBold = false;
            ExcelRange excelRange = ExcelTestBuilder.CreateExcelRange();

            // Act
            excelRange.SetFontBold(expectedToBeBold);

            // Assert
            excelRange.Style.Font.Bold.Should().BeFalse();
        }
    }
}