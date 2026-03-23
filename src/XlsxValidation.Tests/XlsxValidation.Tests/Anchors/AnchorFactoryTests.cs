using FluentAssertions;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;

namespace XlsxValidation.Tests.Anchors;

/// <summary>
/// Тесты для AnchorFactory
/// </summary>
public class AnchorFactoryTests
{
    private readonly AnchorFactory _factory;

    public AnchorFactoryTests()
    {
        _factory = new AnchorFactory();
    }

    [Fact]
    public void Create_ContentAnchor_CreatesCorrectType()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = AnchorType.Content,
            Value = "Тест"
        };

        // Act
        var anchor = _factory.Create(config);

        // Assert
        anchor.GetType().Should().Be(typeof(ContentAnchor));
        anchor.Description.Should().Contain("Тест");
    }

    [Fact]
    public void Create_AddressAnchor_CreatesCorrectType()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = AnchorType.Address,
            Value = "A1"
        };

        // Act
        var anchor = _factory.Create(config);

        // Assert
        anchor.GetType().Should().Be(typeof(AddressAnchor));
        anchor.Description.Should().Contain("A1");
    }

    [Fact]
    public void Create_NamedRangeAnchor_CreatesCorrectType()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = AnchorType.NamedRange,
            Value = "MyRange"
        };

        // Act
        var anchor = _factory.Create(config);

        // Assert
        anchor.GetType().Should().Be(typeof(NamedRangeAnchor));
        anchor.Description.Should().Contain("MyRange");
    }

    [Fact]
    public void Create_OffsetAnchor_WithBaseAnchor_CreatesCorrectType()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = AnchorType.Offset,
            Base = new AnchorConfig { Type = AnchorType.Address, Value = "A1" },
            RowOffset = 1,
            ColOffset = 2
        };

        // Act
        var anchor = _factory.Create(config);

        // Assert
        anchor.GetType().Should().Be(typeof(OffsetAnchor));
        anchor.Description.Should().Contain("Offset");
    }

    [Fact]
    public void Create_OffsetAnchor_WithoutBaseAnchor_ThrowsException()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = AnchorType.Offset,
            RowOffset = 1,
            ColOffset = 2
        };

        // Act & Assert
        _factory.Invoking(f => f.Create(config))
            .Should().Throw<ArgumentException>()
            .WithMessage("*Base anchor required*");
    }

    [Fact]
    public void Create_ContentAnchor_WithoutValue_ThrowsException()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = AnchorType.Content,
            Value = null
        };

        // Act & Assert
        _factory.Invoking(f => f.Create(config))
            .Should().Throw<ArgumentException>()
            .WithMessage("*Value required for content anchor*");
    }

    [Fact]
    public void Create_UnknownType_ThrowsException()
    {
        // Arrange
        var config = new AnchorConfig
        {
            Type = (AnchorType)999,
            Value = "Test"
        };

        // Act & Assert
        _factory.Invoking(f => f.Create(config))
            .Should().Throw<ArgumentException>()
            .WithMessage("*Неизвестный тип якоря*");
    }

    [Fact]
    public void ParseAnchorType_ValidTypes_ReturnsCorrectEnum()
    {
        // Act & Assert
        AnchorFactory.ParseAnchorType("content").Should().Be(AnchorType.Content);
        AnchorFactory.ParseAnchorType("Content").Should().Be(AnchorType.Content);
        AnchorFactory.ParseAnchorType("offset").Should().Be(AnchorType.Offset);
        AnchorFactory.ParseAnchorType("named-range").Should().Be(AnchorType.NamedRange);
        AnchorFactory.ParseAnchorType("address").Should().Be(AnchorType.Address);
    }

    [Fact]
    public void ParseAnchorType_InvalidType_ThrowsException()
    {
        // Act & Assert
        _factory.Invoking(f => AnchorFactory.ParseAnchorType("invalid"))
            .Should().Throw<ArgumentException>();
    }
}
