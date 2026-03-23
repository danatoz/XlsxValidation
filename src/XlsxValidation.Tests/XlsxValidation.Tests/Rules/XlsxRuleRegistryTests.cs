using ClosedXML.Excel;
using FluentAssertions;
using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Results;

namespace XlsxValidation.Tests.Rules;

/// <summary>
/// Тесты для XlsxRuleRegistry
/// </summary>
public class XlsxRuleRegistryTests
{
    private readonly XlsxRuleRegistry _registry;

    public XlsxRuleRegistryTests()
    {
        _registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(_registry);
    }

    [Fact]
    public void RegisterCellRule_NewRule_AddsSuccessfully()
    {
        // Arrange
        var ruleId = "custom-rule";

        // Act & Assert
        _registry.Invoking(r => r.RegisterCellRule(ruleId, _ => (Func<IXLCell, ValidationResult>)(cell => ValidationResult.Ok())))
            .Should().NotThrow();
    }

    [Fact]
    public void RegisterCellRule_DuplicateRule_ThrowsException()
    {
        // Arrange
        var ruleId = "not-empty"; // Уже зарегистрировано во встроенных правилах

        // Act & Assert
        _registry.Invoking(r => r.RegisterCellRule(ruleId, _ => (Func<IXLCell, ValidationResult>)(cell => ValidationResult.Ok())))
            .Should().Throw<InvalidOperationException>()
            .WithMessage($"*'{ruleId}' уже зарегистрировано*");
    }

    [Fact]
    public void RegisterColumnRule_NewRule_AddsSuccessfully()
    {
        // Arrange
        var ruleId = "custom-column-rule";

        // Act & Assert
        _registry.Invoking(r => r.RegisterColumnRule(ruleId, _ => (Func<IXLCell, int, ValidationResult>)((cell, row) => ValidationResult.Ok())))
            .Should().NotThrow();
    }

    [Fact]
    public void RegisterSharedRule_RegistersInBothRegistries()
    {
        // Arrange
        var ruleId = "shared-rule";

        // Act
        _registry.RegisterSharedRule(ruleId, (_, _) => cell => ValidationResult.Ok());

        // Assert
        Assert.NotNull(_registry.GetCellRule(ruleId));
        Assert.NotNull(_registry.GetColumnRule(ruleId));
    }

    [Fact]
    public void GetCellRule_ExistingRule_ReturnsFactory()
    {
        // Act
        var factory = _registry.GetCellRule("not-empty");

        // Assert
        Assert.NotNull(factory);
    }

    [Fact]
    public void GetCellRule_NonExistingRule_ReturnsNull()
    {
        // Act
        var factory = _registry.GetCellRule("non-existing-rule");

        // Assert
        Assert.Null(factory);
    }

    [Fact]
    public void GetColumnRule_ExistingRule_ReturnsFactory()
    {
        // Act
        var factory = _registry.GetColumnRule("not-empty");

        // Assert
        Assert.NotNull(factory);
    }

    [Fact]
    public void IsRegistered_ExistingRule_ReturnsTrue()
    {
        // Act & Assert
        _registry.IsRegistered("not-empty").Should().BeTrue();
    }

    [Fact]
    public void IsRegistered_NonExistingRule_ReturnsFalse()
    {
        // Act & Assert
        _registry.IsRegistered("non-existing").Should().BeFalse();
    }

    [Fact]
    public void GetRegisteredRuleIds_ReturnsAllRegisteredRules()
    {
        // Act
        var ruleIds = _registry.GetRegisteredRuleIds().ToList();

        // Assert
        ruleIds.Should().Contain("not-empty");
        ruleIds.Should().Contain("is-numeric");
        ruleIds.Should().Contain("is-date");
        ruleIds.Should().Contain("max-length");
        ruleIds.Should().Contain("min-value");
        ruleIds.Should().Contain("max-value");
        ruleIds.Should().Contain("matches");
        ruleIds.Should().Contain("one-of");
        ruleIds.Should().Contain("date-not-future");
        ruleIds.Should().Contain("date-not-past");
        ruleIds.Should().Contain("is-merged");
    }

    [Fact]
    public void CellRuleFactory_CreatesExecutableRule()
    {
        // Arrange
        var factory = _registry.GetCellRule("not-empty")!;
        var config = new RuleConfig { Rule = "not-empty" };
        var rule = factory(config);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Test");
        worksheet.Cell("A1").Value = "Тест";

        // Act
        var result = rule(worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void ColumnRuleFactory_CreatesExecutableRule()
    {
        // Arrange
        var factory = _registry.GetColumnRule("not-empty")!;
        var config = new RuleConfig { Rule = "not-empty" };
        var rule = factory(config);

        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Test");
        worksheet.Cell("A1").Value = "Тест";

        // Act
        var result = rule(worksheet.Cell("A1"), 1);

        // Assert
        result.IsValid.Should().BeTrue();
    }
}
