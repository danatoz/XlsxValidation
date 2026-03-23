using ClosedXML.Excel;
using FluentAssertions;
using XlsxValidation.Anchors;
using XlsxValidation.Builder;
using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Validators;

namespace XlsxValidation.Tests.Validators;

/// <summary>
/// Интеграционные тесты для XlsxValidator
/// </summary>
public class XlsxValidatorIntegrationTests : IDisposable
{
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;
    private readonly XlsxRuleRegistry _registry;

    public XlsxValidatorIntegrationTests()
    {
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Данные");
        _registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(_registry);
        SetupTestData();
    }

    private void SetupTestData()
    {
        // Заголовки
        _worksheet.Cell("A1").Value = "Наименование организации";
        _worksheet.Cell("B1").Value = "ООО Ромашка";
        
        _worksheet.Cell("A2").Value = "ИНН";
        _worksheet.Cell("B2").Value = "1234567890";
        
        _worksheet.Cell("A3").Value = "Дата составления";
        _worksheet.Cell("B3").Value = new DateTime(2025, 1, 15);

        // Заголовок таблицы
        _worksheet.Cell("A5").Value = "№";
        _worksheet.Cell("B5").Value = "Наименование";
        _worksheet.Cell("C5").Value = "Количество";
        _worksheet.Cell("D5").Value = "Цена";

        // Данные таблицы
        _worksheet.Cell("A6").Value = 1;
        _worksheet.Cell("B6").Value = "Товар 1";
        _worksheet.Cell("C6").Value = 10;
        _worksheet.Cell("D6").Value = 100.50;

        _worksheet.Cell("A7").Value = 2;
        _worksheet.Cell("B7").Value = "Товар 2";
        _worksheet.Cell("C7").Value = 5;
        _worksheet.Cell("D7").Value = 200.00;
    }

    [Fact]
    public void Validate_ValidData_ReturnsValidReport()
    {
        // Arrange
        var config = new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new WorksheetValidationConfig
                    {
                        Name = "Данные",
                        Cells = new List<CellValidationConfig>
                        {
                            new CellValidationConfig
                            {
                                Name = "Организация",
                                Anchor = new AnchorConfig 
                                { 
                                    Type = AnchorType.Content, 
                                    Value = "Наименование организации" 
                                },
                                Rules = new List<RuleConfig>
                                {
                                    new RuleConfig { Rule = "not-empty" }
                                }
                            }
                        },
                        Tables = new List<TableValidationConfig>
                        {
                            new TableValidationConfig
                            {
                                Name = "Позиции",
                                HeaderAnchor = new AnchorConfig 
                                { 
                                    Type = AnchorType.Content, 
                                    Value = "№" 
                                },
                                Columns = new List<ColumnConfig>
                                {
                                    new ColumnConfig
                                    {
                                        Header = "Наименование",
                                        Rules = new List<RuleConfig>
                                        {
                                            new RuleConfig { Rule = "not-empty" }
                                        }
                                    },
                                    new ColumnConfig
                                    {
                                        Header = "Количество",
                                        Rules = new List<RuleConfig>
                                        {
                                            new RuleConfig { Rule = "is-numeric" },
                                            new RuleConfig { Rule = "min-value", Params = new Dictionary<string, object> { ["min"] = 0 } }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        };

        var validator = new XlsxValidatorBuilder(_registry)
            .WithProfileName("test")
            .FromConfig(config)
            .Build();

        // Act
        var report = validator.Validate(_workbook);

        // Assert
        report.IsValid.Should().BeTrue();
        report.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Validate_InvalidCellData_ReturnsErrors()
    {
        // Arrange: очистим ячейку с организацией
        _worksheet.Cell("B1").Value = string.Empty;

        var config = new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new WorksheetValidationConfig
                    {
                        Name = "Данные",
                        Cells = new List<CellValidationConfig>
                        {
                            new CellValidationConfig
                            {
                                Name = "Организация",
                                Anchor = new AnchorConfig 
                                { 
                                    Type = AnchorType.Offset,
                                    Base = new AnchorConfig { Type = AnchorType.Content, Value = "Наименование организации" },
                                    RowOffset = 0,
                                    ColOffset = 1
                                },
                                Rules = new List<RuleConfig>
                                {
                                    new RuleConfig { Rule = "not-empty" }
                                }
                            }
                        }
                    }
                }
            }
        };

        var validator = new XlsxValidatorBuilder(_registry)
            .WithProfileName("test")
            .FromConfig(config)
            .Build();

        // Act
        var report = validator.Validate(_workbook);

        // Assert
        report.IsValid.Should().BeFalse();
        report.Errors.Should().HaveCount(1);
        report.Errors[0].FieldName.Should().Be("Организация");
        report.Errors[0].Message.Should().Contain("пустой");
    }

    [Fact]
    public void Validate_InvalidTableData_ReturnsErrors()
    {
        // Arrange: установим нечисловое значение в колонку Количество
        _worksheet.Cell("C6").Value = "не число";

        var config = new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new WorksheetValidationConfig
                    {
                        Name = "Данные",
                        Tables = new List<TableValidationConfig>
                        {
                            new TableValidationConfig
                            {
                                Name = "Позиции",
                                HeaderAnchor = new AnchorConfig 
                                { 
                                    Type = AnchorType.Content, 
                                    Value = "№" 
                                },
                                Columns = new List<ColumnConfig>
                                {
                                    new ColumnConfig
                                    {
                                        Header = "Количество",
                                        Rules = new List<RuleConfig>
                                        {
                                            new RuleConfig { Rule = "is-numeric" }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        };

        var validator = new XlsxValidatorBuilder(_registry)
            .WithProfileName("test")
            .FromConfig(config)
            .Build();

        // Act
        var report = validator.Validate(_workbook);

        // Assert
        report.IsValid.Should().BeFalse();
        report.Errors.Should().HaveCount(1);
        report.Errors[0].RowNumber.Should().Be(6);
        report.Errors[0].Message.Should().Contain("числом");
    }

    [Fact]
    public void Validate_AnchorNotFound_ReturnsError()
    {
        // Arrange
        var config = new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new WorksheetValidationConfig
                    {
                        Name = "Данные",
                        Cells = new List<CellValidationConfig>
                        {
                            new CellValidationConfig
                            {
                                Name = "Несуществующее",
                                Anchor = new AnchorConfig 
                                { 
                                    Type = AnchorType.Content, 
                                    Value = "Несуществующее значение" 
                                },
                                Rules = new List<RuleConfig>
                                {
                                    new RuleConfig { Rule = "not-empty" }
                                }
                            }
                        }
                    }
                }
            }
        };

        var validator = new XlsxValidatorBuilder(_registry)
            .WithProfileName("test")
            .FromConfig(config)
            .Build();

        // Act
        var report = validator.Validate(_workbook);

        // Assert
        report.IsValid.Should().BeFalse();
        report.Errors.Should().HaveCount(1);
        report.Errors[0].RuleId.Should().Be("AnchorNotFound");
    }

    [Fact]
    public void Validate_WithStream_WorksCorrectly()
    {
        // Arrange
        using var stream = new MemoryStream();
        _workbook.SaveAs(stream);
        stream.Position = 0;

        var config = new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new WorksheetValidationConfig
                    {
                        Name = "Данные",
                        Cells = new List<CellValidationConfig>
                        {
                            new CellValidationConfig
                            {
                                Name = "Организация",
                                Anchor = new AnchorConfig 
                                { 
                                    Type = AnchorType.Content, 
                                    Value = "Наименование организации" 
                                },
                                Rules = new List<RuleConfig>()
                            }
                        }
                    }
                }
            }
        };

        var validator = new XlsxValidatorBuilder(_registry)
            .WithProfileName("test")
            .FromConfig(config)
            .Build();

        // Act
        var report = validator.Validate(stream);

        // Assert
        report.ProfileName.Should().Be("test");
    }

    [Fact]
    public void ValidateAndThrow_InvalidData_ThrowsXlsxValidationException()
    {
        // Arrange: очистим ячейку с организацией (B1)
        _worksheet.Cell("B1").Value = string.Empty;

        var config = new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new WorksheetValidationConfig
                    {
                        Name = "Данные",
                        Cells = new List<CellValidationConfig>
                        {
                            new CellValidationConfig
                            {
                                Name = "Организация",
                                Anchor = new AnchorConfig
                                {
                                    Type = AnchorType.Offset,
                                    Base = new AnchorConfig
                                    {
                                        Type = AnchorType.Content,
                                        Value = "Наименование организации"
                                    },
                                    RowOffset = 0,
                                    ColOffset = 1
                                },
                                Rules = new List<RuleConfig>
                                {
                                    new RuleConfig { Rule = "not-empty" }
                                }
                            }
                        }
                    }
                }
            }
        };

        var validator = new XlsxValidatorBuilder(_registry)
            .WithProfileName("test")
            .FromConfig(config)
            .Build();

        // Act & Assert
        validator.Invoking(v => v.ValidateAndThrow(_workbook))
            .Should().Throw<XlsxValidationException>()
            .Where(ex => ex.Report.IsValid == false);
    }

    public void Dispose()
    {
        _workbook.Dispose();
    }
}
