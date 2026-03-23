using ClosedXML.Excel;
using FluentAssertions;
using XlsxValidation.Anchors;

namespace XlsxValidation.Tests.Anchors;

/// <summary>
/// Тесты для OffsetAnchor
/// </summary>
public class OffsetAnchorTests : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;

    public OffsetAnchorTests()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("TestSheet");
        SetupTestData();
    }

    private void SetupTestData()
    {
        _worksheet.Cell("A1").Value = "Заголовок";
        _worksheet.Cell("B1").Value = "Значение 1";
        _worksheet.Cell("C1").Value = "Значение 2";
        _worksheet.Cell("A2").Value = "Строка 2";
        _worksheet.Cell("B2").Value = 123;
    }

    [Fact]
    public void Resolve_ColumnOffset_ReturnsCorrectCell()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Заголовок");
        var offsetAnchor = new OffsetAnchor(baseAnchor, 0, 1);

        // Act
        var result = offsetAnchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("B1");
        result.Cell!.GetValue<string>().Should().Be("Значение 1");
    }

    [Fact]
    public void Resolve_RowOffset_ReturnsCorrectCell()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Заголовок");
        var offsetAnchor = new OffsetAnchor(baseAnchor, 1, 0);

        // Act
        var result = offsetAnchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("A2");
    }

    [Fact]
    public void Resolve_RowAndColumnOffset_ReturnsCorrectCell()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Заголовок");
        var offsetAnchor = new OffsetAnchor(baseAnchor, 1, 1);

        // Act
        var result = offsetAnchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("B2");
        result.Cell!.GetValue<int>().Should().Be(123);
    }

    [Fact]
    public void Resolve_NegativeOffset_ReturnsCorrectCell()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Значение 2");
        var offsetAnchor = new OffsetAnchor(baseAnchor, 0, -2);

        // Act
        var result = offsetAnchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("A1");
    }

    [Fact]
    public void Resolve_BaseAnchorNotFound_ReturnsFailure()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Несуществующее");
        var offsetAnchor = new OffsetAnchor(baseAnchor, 1, 1);

        // Act
        var result = offsetAnchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeFalse();
        result.ErrorMessage.Should().Contain("Базовый якорь не найден");
    }

    [Fact]
    public void Resolve_ChainedOffsets_WorksCorrectly()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Заголовок");
        var firstOffset = new OffsetAnchor(baseAnchor, 0, 1); // B1
        var secondOffset = new OffsetAnchor(firstOffset, 1, 1); // C2

        // Act
        var result = secondOffset.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("C2");
    }

    [Fact]
    public void Description_ReturnsCorrectFormat()
    {
        // Arrange
        var baseAnchor = new ContentAnchor("Тест");
        var offsetAnchor = new OffsetAnchor(baseAnchor, 2, -1);

        // Act & Assert
        offsetAnchor.Description.Should().Be("Offset от Content: 'Тест': row=+2, col=-1");
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}
