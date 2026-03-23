using ClosedXML.Excel;
using FluentAssertions;
using XlsxValidation.Anchors;

namespace XlsxValidation.Tests.Anchors;

/// <summary>
/// Тесты для AddressAnchor
/// </summary>
public class AddressAnchorTests : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;

    public AddressAnchorTests()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("TestSheet");
        SetupTestData();
    }

    private void SetupTestData()
    {
        _worksheet.Cell("A1").Value = "Ячейка A1";
        _worksheet.Cell("B5").Value = 42;
        _worksheet.Cell("C10").Value = new DateTime(2025, 3, 15);
    }

    [Fact]
    public void Resolve_ValidAddress_ReturnsCorrectCell()
    {
        // Arrange
        var anchor = new AddressAnchor("B5");

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.Address.ToString().Should().Be("B5");
        result.Cell!.GetValue<int>().Should().Be(42);
    }

    [Fact]
    public void Resolve_DifferentAddress_ReturnsCorrectCell()
    {
        // Arrange
        var anchor = new AddressAnchor("C10");

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeTrue();
        result.Cell!.GetValue<DateTime>().Should().Be(new DateTime(2025, 3, 15));
    }

    [Fact]
    public void Resolve_InvalidAddress_ReturnsFailure()
    {
        // Arrange
        var anchor = new AddressAnchor("INVALID");

        // Act
        var result = anchor.Resolve(_worksheet);

        // Assert
        result.IsSuccess.Should().BeFalse();
        result.ErrorMessage.Should().Contain("Некорректный адрес");
    }

    [Fact]
    public void Description_ReturnsCorrectFormat()
    {
        // Arrange
        var anchor = new AddressAnchor("D15");

        // Act & Assert
        anchor.Description.Should().Be("Address: 'D15'");
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}
