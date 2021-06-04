using System.Drawing;
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
    public class FontTests
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
        public void SetFontBold_WithValidValue_ReturnsInstanceOfSut()
        {
            // Arrange
            bool value = _fixture.Create<bool>();

            // Act
            ISpreadsheetWriter result = _sut.SetFontBold(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void SetFontBold_WithActive_MarksExcelRangeAsBold()
        {
            // Arrange
            bool input = true;

            // Act
            _sut.SetFontBold(input);

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Font.Bold.Should().BeTrue();
        }

        [TestMethod]
        public void SetFontBold_WithInactive_MarksExcelRangeAsBold()
        {
            // Arrange
            bool input = false;

            // Act
            _sut.SetFontBold(input);

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Font.Bold.Should().BeFalse();
        }

        [TestMethod]
        public void SetBackgroundColor_WithValidColor_ReturnsInstanceOfSut()
        {
            // Arrange
            Color value = _fixture.Create<Color>();

            // Act
            ISpreadsheetWriter result = _sut.SetBackgroundColor(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void SetFontColor_WithValidColor_ReturnsInstanceOfSut()
        {
            // Arrange
            Color value = _fixture.Create<Color>();

            // Act
            ISpreadsheetWriter result = _sut.SetFontColor(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void SetFontSize_WithValidSize_ReturnsInstanceOfSut()
        {
            // Arrange
            float value = _fixture.Create<float>();

            // Act
            ISpreadsheetWriter result = _sut.SetFontSize(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void SetFontSize_WithValidSize_SetsFontSizeOfExcelRange()
        {
            // Arrange
            float value = _fixture.Create<float>();

            // Act
            _sut.SetFontSize(value);

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Font.Size.Should().Be(value);
        }

        [TestMethod]
        public void SetFormat_WithValidFormat_ReturnsInstanceOfSut()
        {
            // Arrange
            string value = _fixture.Create<string>();

            // Act
            ISpreadsheetWriter result = _sut.SetFormat(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void SetFormat_WithValidFormat_SetsFormatOfExcelRange()
        {
            // Arrange
            string value = _fixture.Create<string>();

            // Act
            _sut.SetFormat(value);

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Numberformat.Format.Should().Be(value);
        }

        [TestMethod]
        public void SetTextRotation_WithValidRotation_ReturnsInstanceOfSut()
        {
            // Arrange
            int value = 10;

            // Act
            ISpreadsheetWriter result = _sut.SetTextRotation(value);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void SetTextRotation_WithValidRotation_SetsRotationOfExcelRange()
        {
            // Arrange
            int value = 10;

            // Act
            _sut.SetTextRotation(value);

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.TextRotation.Should().Be(value);
        }

        [TestMethod]
        public void ResetStyling_WithValidWriter_ReturnsInstanceOfSut()
        {
            // Arrange
            // Act
            ISpreadsheetWriter result = _sut.ResetStyling();

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void ResetStyling_WithConfiguredFontSize_ResetsFontSizeOfExcelRange()
        {
            // Arrange
            float value = _fixture.Create<float>();
            _sut.SetFontSize(value);

            // Act
            _sut.ResetStyling();

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Font.Size.Should().Be(11);
        }

        [TestMethod]
        public void ResetStyling_WithConfiguredFontBold_ResetsFontBoldOfExcelRange()
        {
            // Arrange
            _sut.SetFontBold(true);

            // Act
            _sut.ResetStyling();

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Font.Bold.Should().BeFalse();
        }

        [TestMethod]
        public void ResetStyling_WithConfiguredFormat_ResetsFormatOfExcelRange()
        {
            // Arrange
            string randomFormat = _fixture.Create<string>();
            _sut.SetFormat(randomFormat);

            // Act
            _sut.ResetStyling();

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.Numberformat.Format.Should().Be("General");
        }

        [TestMethod]
        public void ResetStyling_WithConfiguredTextRotation_ResetsTextRotationOfExcelRange()
        {
            // Arrange
            int randomRotation = 46;
            _sut.SetTextRotation(randomRotation);

            // Act
            _sut.ResetStyling();

            // Assert
            _sut.Write(string.Empty);
            var cell = _worksheet.GetCell(_sut.CurrentPosition);
            cell.Style.TextRotation.Should().Be(0);
        }
    }
}
