using AutoFixture;
using FluentAssertions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using SpreadsheetWriter.EPPlus.UnitTests.Builders;

namespace SpreadsheetWriter.EPPlus.UnitTests.ExcelSpreadsheetWriterTests
{
    [TestClass]
    public class MoveTests
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
        public void MoveDown_WithValidSpreadsheet_ReturnsInstanceOfSut()
        {
            // Arrange

            // Act
            var result = _sut.MoveDown();

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveDown_WithValidSpreadsheet_IncreasesYAxisValue()
        {
            // Arrange

            // Act
            _sut.MoveDown();

            // Assert
            _sut.CurrentPosition.X.Should().Be(1);
            _sut.CurrentPosition.Y.Should().Be(2);
        }

        [TestMethod]
        public void MoveUp_WithValidSpreadsheet_ReturnsInstanceOfSut()
        {
            // Arrange

            // Act
            var result = _sut.MoveUp();

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveUp_WithValidSpreadsheet_DecreasesYAxisValue()
        {
            // Arrange

            // Act
            _sut.MoveUp();

            // Assert
            _sut.CurrentPosition.X.Should().Be(1);
            _sut.CurrentPosition.Y.Should().Be(0);
        }

        [TestMethod]
        public void MoveLeft_WithValidSpreadsheet_ReturnsInstanceOfSut()
        {
            // Arrange

            // Act
            var result = _sut.MoveLeft();

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveLeft_WithValidSpreadsheet_DecreasesXAxisValue()
        {
            // Arrange

            // Act
            _sut.MoveLeft();

            // Assert
            _sut.CurrentPosition.X.Should().Be(0);
            _sut.CurrentPosition.Y.Should().Be(1);
        }

        [TestMethod]
        public void MoveRight_WithValidSpreadsheet_ReturnsInstanceOfSut()
        {
            // Arrange

            // Act
            var result = _sut.MoveRight();

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveRight_WithValidSpreadsheet_IncreasesXAxisValue()
        {
            // Arrange

            // Act
            _sut.MoveRight();

            // Assert
            _sut.CurrentPosition.X.Should().Be(2);
            _sut.CurrentPosition.Y.Should().Be(1);
        }

        [TestMethod]
        public void MoveDownTimes_WithValidNumberOfTimes_ReturnsInstanceOfSut()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            var result = _sut.MoveDownTimes(times);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveDownTimes_WithValidNumberOfTimes_IncreasesYAxisValueNumberOfTimes()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            _sut.MoveDownTimes(times);

            // Assert
            _sut.CurrentPosition.X.Should().Be(1);
            _sut.CurrentPosition.Y.Should().Be(times + 1);
        }

        [TestMethod]
        public void MoveUpTimes_WithValidNumberOfTimes_ReturnsInstanceOfSut()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            var result = _sut.MoveUpTimes(times);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveUpTimes_WithValidNumberOfTimes_DecreasesYAxisValueNumberOfTimes()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            _sut.MoveUpTimes(times);

            // Assert
            _sut.CurrentPosition.X.Should().Be(1);
            _sut.CurrentPosition.Y.Should().Be(-times + 1);
        }

        [TestMethod]
        public void MoveLeftTimes_WithValidNumberOfTimes_ReturnsInstanceOfSut()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            var result = _sut.MoveLeftTimes(times);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveLeftTimes_WithValidNumberOfTimes_DecreasesXAxisValueNumberOfTimes()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            _sut.MoveLeftTimes(times);

            // Assert
            _sut.CurrentPosition.X.Should().Be(-times + 1);
            _sut.CurrentPosition.Y.Should().Be(1);
        }

        [TestMethod]
        public void MoveRightTimes_WithValidNumberOfTimes_ReturnsInstanceOfSut()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            var result = _sut.MoveLeftTimes(times);

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void MoveRightTimes_WithValidNumberOfTimes_IncreasesXAxisValueNumberOfTimes()
        {
            // Arrange
            var times = _fixture.Create<int>();

            // Act
            _sut.MoveRightTimes(times);

            // Assert
            _sut.CurrentPosition.X.Should().Be(times + 1);
            _sut.CurrentPosition.Y.Should().Be(1);
        }

        [TestMethod]
        public void NewLine_WithValidWriter_ReturnsInstanceOfSut()
        {
            // Arrange
            // Act
            var result = _sut.NewLine();

            // Assert
            result.Should().Be(_sut);
        }

        [TestMethod]
        public void NewLine_WithValidWriter_MovesOneRowDownAndAllTheWayLeft()
        {
            // Arrange
            // Act
            _sut.NewLine();

            // Assert
            _sut.CurrentPosition.X.Should().Be(1);
            _sut.CurrentPosition.Y.Should().Be(2);
        }
    }
}