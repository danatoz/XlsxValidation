using ClosedXML.Excel;
using FluentAssertions;
using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Results;

namespace XlsxValidation.Tests.Rules;

/// <summary>
/// Тесты для встроенных правил валидации
/// </summary>
public class BuiltInRulesTests : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;

    public BuiltInRulesTests()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("TestSheet");
    }

    [Fact]
    public void NotEmpty_EmptyCell_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = string.Empty;
        var rule = CreateRule("not-empty");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain("пустой");
    }

    [Fact]
    public void NotEmpty_NonEmptyCell_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "Тест";
        var rule = CreateRule("not-empty");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void IsNumeric_NumericValue_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = 123.45;
        var rule = CreateRule("is-numeric");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void IsNumeric_TextValue_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "Не число";
        var rule = CreateRule("is-numeric");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain("числом");
    }

    [Fact]
    public void IsDate_DateValue_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = new DateTime(2025, 1, 15);
        var rule = CreateRule("is-date");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void IsDate_TextValue_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "Не дата";
        var rule = CreateRule("is-date");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain("датой");
    }

    [Fact]
    public void MaxLength_ExceedsMax_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "Очень длинная строка";
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "max-length",
            Params = new Dictionary<string, object> { ["max"] = 5 }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("max-length")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain("5");
    }

    [Fact]
    public void MaxLength_WithinLimit_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "Тест";
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "max-length",
            Params = new Dictionary<string, object> { ["max"] = 10 }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("max-length")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void MinValue_BelowMin_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = 5;
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "min-value",
            Params = new Dictionary<string, object> { ["min"] = 10 }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("min-value")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain(">= 10");
    }

    [Fact]
    public void MaxValue_AboveMax_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = 100;
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "max-value",
            Params = new Dictionary<string, object> { ["max"] = 50 }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("max-value")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain("<= 50");
    }

    [Fact]
    public void Matches_PatternMatch_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "123456";
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "matches",
            Params = new Dictionary<string, object> { ["pattern"] = @"^\d{6}$" }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("matches")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void Matches_PatternMismatch_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "abc123";
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "matches",
            Params = new Dictionary<string, object> { ["pattern"] = @"^\d{6}$" }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("matches")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
    }

    [Fact]
    public void OneOf_ValidValue_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "шт";
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "one-of",
            Params = new Dictionary<string, object> 
            { 
                ["values"] = new List<object> { "шт", "кг", "л", "м" }
            }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("one-of")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void OneOf_InvalidValue_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = "фут";
        var config = new XlsxValidation.Configuration.RuleConfig
        {
            Rule = "one-of",
            Params = new Dictionary<string, object> 
            { 
                ["values"] = new List<object> { "шт", "кг", "л", "м" }
            }
        };
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule("one-of")!;
        var rule = factory(config);

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
    }

    [Fact]
    public void DateNotFuture_FutureDate_ReturnsError()
    {
        // Arrange
        _worksheet.Cell("A1").Value = DateTime.Today.AddDays(1);
        var rule = CreateRule("date-not-future");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeFalse();
        result.ErrorMessage.Should().Contain("будущем");
    }

    [Fact]
    public void DateNotFuture_PastDate_ReturnsOk()
    {
        // Arrange
        _worksheet.Cell("A1").Value = DateTime.Today.AddDays(-1);
        var rule = CreateRule("date-not-future");

        // Act
        var result = rule(_worksheet.Cell("A1"));

        // Assert
        result.IsValid.Should().BeTrue();
    }

    private Func<IXLCell, ValidationResult> CreateRule(string ruleId)
    {
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        var factory = registry.GetCellRule(ruleId)!;
        return factory(new XlsxValidation.Configuration.RuleConfig { Rule = ruleId });
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}
