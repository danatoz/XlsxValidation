using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;
using XlsxValidation.Parsing;

namespace XlsxValidation.Tests.Parsing;

public class XlsxParserIntegrationTests : IDisposable
{
    private readonly string _testFilePath;
    private readonly XLWorkbook _workbook;
    private readonly IXLWorksheet _worksheet;

    public XlsxParserIntegrationTests()
    {
        _testFilePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
        _workbook = new XLWorkbook();
        _worksheet = _workbook.AddWorksheet("Data");

        SetupTestData();
    }

    private void SetupTestData()
    {
        // Заголовки
        _worksheet.Cell(1, 1).Value = "Organization";
        _worksheet.Cell(1, 2).Value = "INN";
        _worksheet.Cell(1, 3).Value = "Date";
        _worksheet.Cell(1, 4).Value = "Amount";

        // Данные
        _worksheet.Cell(2, 1).Value = "Test Organization";
        _worksheet.Cell(2, 2).Value = "1234567890";
        _worksheet.Cell(2, 3).Value = new DateTime(2024, 1, 15);
        _worksheet.Cell(2, 4).Value = 1000.50m;

        _worksheet.Cell(3, 1).Value = "Another Org";
        _worksheet.Cell(3, 2).Value = "0987654321";
        _worksheet.Cell(3, 3).Value = new DateTime(2024, 2, 20);
        _worksheet.Cell(3, 4).Value = 2500.75m;

        _workbook.SaveAs(_testFilePath);
    }

    [Fact]
    public void Parse_File_Returns_Success()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        // Act
        var result = parser.Parse(_testFilePath);

        // Assert
        Assert.True(result.IsSuccess);
        Assert.NotEmpty(result.Fields);
        Assert.NotEmpty(result.Tables);
    }

    [Fact]
    public void Parse_File_Extracts_Fields()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        // Act
        var result = parser.Parse(_testFilePath);

        // Assert
        var orgField = result.GetField("Organization");
        Assert.NotNull(orgField);
        Assert.Equal("Test Organization", orgField.RawValue);

        var innField = result.GetField("INN");
        Assert.NotNull(innField);
        Assert.Equal("1234567890", innField.RawValue);
    }

    [Fact]
    public void Parse_File_Extracts_Tables()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        // Act
        var result = parser.Parse(_testFilePath);

        // Assert
        var table = result.GetTable("Items");
        Assert.NotNull(table);
        Assert.Equal(2, table.RowCount);
        Assert.Contains("Organization", table.Headers);
        Assert.Contains("INN", table.Headers);
    }

    [Fact]
    public void Parse_Stream_Works_Correctly()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        using var stream = new FileStream(_testFilePath, FileMode.Open, FileAccess.Read);

        // Act
        var result = parser.Parse(stream);

        // Assert
        Assert.True(result.IsSuccess);
    }

    [Fact]
    public void Parse_Workbook_Works_Correctly()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        // Act
        var result = parser.Parse(_workbook);

        // Assert
        Assert.True(result.IsSuccess);
    }

    [Fact]
    public void ParsedField_AsType_Works_Correctly()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        // Act
        var result = parser.Parse(_testFilePath);

        // Assert
        var amountField = result.Fields.FirstOrDefault(f => f.Name == "Amount");
        Assert.NotNull(amountField);

        var amount = amountField.AsDecimal();
        Assert.Equal(1000.50m, amount);

        var date = result.Fields.FirstOrDefault(f => f.Name == "Date");
        Assert.NotNull(date);
        var dateValue = date.AsDateTime();
        Assert.Equal(new DateTime(2024, 1, 15), dateValue);
    }

    [Fact]
    public void MapTo_Maps_To_Model()
    {
        // Arrange
        var config = CreateTestConfig();
        var typeConverter = new TypeConverter(config.Parsing.Options);
        var parser = XlsxParser.FromConfig("test", config, typeConverter);

        // Act
        var result = parser.Parse(_testFilePath);
        var model = result.MapTo<TestInvoiceModel>();

        // Assert
        Assert.Equal("Test Organization", model.Organization);
        Assert.Equal("1234567890", model.Inn);
    }

    private XlsxProfileConfig CreateTestConfig()
    {
        return new XlsxProfileConfig
        {
            Profile = "test",
            Validation = new ProfileValidationSection
            {
                Worksheets = new List<WorksheetValidationConfig>
                {
                    new()
                    {
                        Name = "Data",
                        Cells = new List<CellValidationConfig>
                        {
                            new()
                            {
                                Name = "Organization",
                                Anchor = new AnchorConfig
                                {
                                    Type = AnchorType.Address,
                                    Value = "A2"
                                },
                                Rules = new List<RuleConfig>()
                            },
                            new()
                            {
                                Name = "INN",
                                Anchor = new AnchorConfig
                                {
                                    Type = AnchorType.Address,
                                    Value = "B2"
                                },
                                Rules = new List<RuleConfig>()
                            },
                            new()
                            {
                                Name = "Date",
                                Anchor = new AnchorConfig
                                {
                                    Type = AnchorType.Address,
                                    Value = "C2"
                                },
                                Rules = new List<RuleConfig>()
                            },
                            new()
                            {
                                Name = "Amount",
                                Anchor = new AnchorConfig
                                {
                                    Type = AnchorType.Address,
                                    Value = "D2"
                                },
                                Rules = new List<RuleConfig>()
                            }
                        },
                        Tables = new List<TableValidationConfig>
                        {
                            new()
                            {
                                Name = "Items",
                                HeaderAnchor = new AnchorConfig
                                {
                                    Type = AnchorType.Address,
                                    Value = "A1"
                                },
                                StopCondition = new StopConditionConfig
                                {
                                    Type = StopConditionType.EmptyRow
                                },
                                Columns = new List<ColumnConfig>
                                {
                                    new() { Header = "Organization" },
                                    new() { Header = "INN" },
                                    new() { Header = "Date" },
                                    new() { Header = "Amount" }
                                }
                            }
                        }
                    }
                }
            },
            Parsing = new ParsingSection
            {
                FieldTypes = new Dictionary<string, string>
                {
                    { "Organization", "string" },
                    { "INN", "string" },
                    { "Date", "date" },
                    { "Amount", "decimal" }
                },
                Options = new ParseOptions
                {
                    Culture = "ru-RU",
                    TrimStrings = true,
                    DateFormats = new[] { "dd.MM.yyyy" },
                    SkipEmptyCells = true
                }
            }
        };
    }

    public void Dispose()
    {
        _workbook.Dispose();

        if (File.Exists(_testFilePath))
        {
            try
            {
                File.Delete(_testFilePath);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }

    // Test model
    private class TestInvoiceModel
    {
        [XlsxField(Name = "Organization")]
        public string Organization { get; set; } = string.Empty;

        [XlsxField(Name = "INN")]
        public string Inn { get; set; } = string.Empty;
    }
}
