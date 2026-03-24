using XlsxValidation.Parsing;

namespace XlsxValidation.Tests.Parsing;

public class XlsxParseResultTests
{
    public class IsSuccess
    {
        [Fact]
        public void Returns_True_When_No_Errors()
        {
            var result = XlsxParseResult.Empty("test");
            Assert.True(result.IsSuccess);
        }

        [Fact]
        public void Returns_False_When_Has_Errors()
        {
            var errors = new List<ParseError>
            {
                ParseError.Create("field", "error message")
            };
            var result = XlsxParseResult.WithErrors("test", errors);
            Assert.False(result.IsSuccess);
        }
    }

    public class GetField
    {
        [Fact]
        public void Returns_Field_By_Name()
        {
            var fields = new List<ParsedField>
            {
                new() { Name = "Field1", RawValue = "Value1" },
                new() { Name = "Field2", RawValue = "Value2" }
            };

            var result = XlsxParseResult.Success("test", fields, new List<ParsedTable>());

            var field = result.GetField("Field1");

            Assert.NotNull(field);
            Assert.Equal("Value1", field.RawValue);
        }

        [Fact]
        public void Returns_Null_For_Unknown_Field()
        {
            var result = XlsxParseResult.Empty("test");
            var field = result.GetField("Unknown");

            Assert.Null(field);
        }
    }

    public class GetTable
    {
        [Fact]
        public void Returns_Table_By_Name()
        {
            var tables = new List<ParsedTable>
            {
                new() { Name = "Table1", Headers = new List<string>(), Rows = new List<ParsedTableRow>() },
                new() { Name = "Table2", Headers = new List<string>(), Rows = new List<ParsedTableRow>() }
            };

            var result = XlsxParseResult.Success("test", new List<ParsedField>(), tables);

            var table = result.GetTable("Table1");

            Assert.NotNull(table);
            Assert.Equal("Table1", table.Name);
        }

        [Fact]
        public void Returns_Null_For_Unknown_Table()
        {
            var result = XlsxParseResult.Empty("test");
            var table = result.GetTable("Unknown");

            Assert.Null(table);
        }
    }

    public class Empty
    {
        [Fact]
        public void Creates_Empty_Result_With_Profile_Name()
        {
            var result = XlsxParseResult.Empty("my-profile");

            Assert.Equal("my-profile", result.ProfileName);
            Assert.Empty(result.Fields);
            Assert.Empty(result.Tables);
            Assert.Empty(result.Errors);
            Assert.True(result.IsSuccess);
        }
    }

    public class WithErrors
    {
        [Fact]
        public void Creates_Result_With_Errors()
        {
            var errors = new List<ParseError>
            {
                ParseError.Create("field1", "error 1"),
                ParseError.Create("field2", "error 2")
            };

            var result = XlsxParseResult.WithErrors("test", errors);

            Assert.Equal("test", result.ProfileName);
            Assert.Equal(2, result.Errors.Count);
            Assert.False(result.IsSuccess);
        }
    }

    public class Success
    {
        [Fact]
        public void Creates_Successful_Result_With_Fields_And_Tables()
        {
            var fields = new List<ParsedField>
            {
                new() { Name = "Field1", RawValue = "Value1" }
            };

            var tables = new List<ParsedTable>
            {
                new() { Name = "Table1", Headers = new List<string> { "H1" }, Rows = new List<ParsedTableRow>() }
            };

            var result = XlsxParseResult.Success("test", fields, tables);

            Assert.Equal("test", result.ProfileName);
            Assert.Single(result.Fields);
            Assert.Single(result.Tables);
            Assert.Empty(result.Errors);
            Assert.True(result.IsSuccess);
        }
    }

    public class ParseErrorTests
    {
        [Fact]
        public void Create_Creates_Error_With_All_Properties()
        {
            var error = ParseError.Create(
                fieldName: "TestField",
                message: "Test error message",
                cellAddress: "A1",
                rowNumber: 5,
                worksheetName: "Sheet1",
                exception: new InvalidOperationException("test"));

            Assert.Equal("TestField", error.FieldName);
            Assert.Equal("Test error message", error.Message);
            Assert.Equal("A1", error.CellAddress);
            Assert.Equal(5, error.RowNumber);
            Assert.Equal("Sheet1", error.WorksheetName);
            Assert.NotNull(error.Exception);
        }

        [Fact]
        public void Create_Creates_Error_With_Minimal_Properties()
        {
            var error = ParseError.Create("Field", "Message");

            Assert.Equal("Field", error.FieldName);
            Assert.Equal("Message", error.Message);
            Assert.Null(error.CellAddress);
            Assert.Null(error.RowNumber);
            Assert.Null(error.WorksheetName);
            Assert.Null(error.Exception);
        }
    }

    public class ParsedFieldTests
    {
        [Fact]
        public void IsEmpty_Returns_True_For_Null_Value()
        {
            var field = new ParsedField { RawValue = null };
            Assert.True(field.IsEmpty);
        }

        [Fact]
        public void IsEmpty_Returns_True_For_Empty_String()
        {
            var field = new ParsedField { RawValue = "" };
            Assert.True(field.IsEmpty);
        }

        [Fact]
        public void IsEmpty_Returns_True_For_Whitespace_String()
        {
            var field = new ParsedField { RawValue = "   " };
            Assert.True(field.IsEmpty);
        }

        [Fact]
        public void IsEmpty_Returns_False_For_Value()
        {
            var field = new ParsedField { RawValue = "value" };
            Assert.False(field.IsEmpty);
        }
    }

    public class ParsedTableTests
    {
        [Fact]
        public void RowCount_Returns_Number_Of_Rows()
        {
            var rows = new List<ParsedTableRow>
            {
                new() { RowNumber = 1 },
                new() { RowNumber = 2 },
                new() { RowNumber = 3 }
            };

            var table = new ParsedTable
            {
                Name = "Test",
                Headers = new List<string>(),
                Rows = rows
            };

            Assert.Equal(3, table.RowCount);
        }

        [Fact]
        public void GetRow_Returns_Row_By_Index()
        {
            var rows = new List<ParsedTableRow>
            {
                new() { RowNumber = 1, Fields = new() },
                new() { RowNumber = 2, Fields = new() },
                new() { RowNumber = 3, Fields = new() }
            };

            var table = new ParsedTable
            {
                Name = "Test",
                Headers = new List<string>(),
                Rows = rows
            };

            var row = table.GetRow(1);

            Assert.NotNull(row);
            Assert.Equal(2, row.RowNumber);
        }

        [Fact]
        public void GetRow_Returns_Null_For_Invalid_Index()
        {
            var table = new ParsedTable
            {
                Name = "Test",
                Headers = new List<string>(),
                Rows = new List<ParsedTableRow>()
            };

            Assert.Null(table.GetRow(-1));
            Assert.Null(table.GetRow(0));
            Assert.Null(table.GetRow(100));
        }
    }
}
