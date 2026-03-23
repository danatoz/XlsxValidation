using ClosedXML.Excel;
using FluentAssertions;
using XlsxValidation.Anchors;

namespace XlsxValidation.Tests.Anchors;

/// <summary>
/// Тесты для ContentAnchor
/// </summary>
public class ContentAnchorTests : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;

    public ContentAnchorTests()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("TestSheet");
        SetupTestData();
    }

    private void SetupTestData()
    {
        _worksheet.Cell("A1").Value = "Наименование организации";
        _worksheet.Cell("B1").Value = "ООО Ромашка";
        _worksheet.Cell("A2").Value = "ИНН";
        _worksheet.Cell("B2").Value = "1234567890";
        _worksheet.Cell("A3").Value = "Дата составления";
        _worksheet.Cell("B3").Value = new DateTime(2025, 1, 15);
    }

    [Fact]
    public void Resolve_ExistingContent_FindsCell()
    {
        // Arrange
        var anchor = new ContentAnchor("ИНН");

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("A2");
    }

    [Fact]
    public void Resolve_NonExistingContent_ReturnsFailure()
    {
        // Arrange
        var anchor = new ContentAnchor("Несуществующее значение");

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeFalse();
        result.ErrorMessage.Should().Contain("не найдена");
    }

    [Fact]
    public void Resolve_PartialMatch_FindsCell()
    {
        // Arrange
        var anchor = new ContentAnchor("организации", exactMatch: false);

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("A1");
    }

    [Fact]
    public void Resolve_ExactMatch_NoMatchForPartial()
    {
        // Arrange
        var anchor = new ContentAnchor("организации", exactMatch: true);

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeFalse();
    }

    [Fact]
    public void Resolve_ExactMatch_FindsExactCell()
    {
        // Arrange
        var anchor = new ContentAnchor("ИНН", exactMatch: true);

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("A2");
    }

    [Fact]
    public void Description_ReturnsCorrectFormat()
    {
        // Arrange
        var anchor = new ContentAnchor("Тест");

        // Act & Assert
        anchor.Description.Should().Be("Content: 'Тест'");
    }

    [Fact]
    public void Description_WithExactMatch_IncludesExactText()
    {
        // Arrange
        var anchor = new ContentAnchor("Тест", exactMatch: true);

        // Act & Assert
        anchor.Description.Should().Be("Content: 'Тест' (точное)");
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}
