using ClosedXML.Excel;
using XlsxValidation.Parsing;

namespace XlsxValidation.Tests.Parsing;

public class ParsedFieldExtensionsTests
{
    private static ParsedField CreateField(string? value, XLDataType dataType = XLDataType.Text)
    {
        return new ParsedField
        {
            Name = "TestField",
            RawValue = value,
            DataType = dataType,
            CellAddress = "A1",
            WorksheetName = "Sheet1"
        };
    }

    public class AsString : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Returns_String_Value()
        {
            var field = CreateField("  Hello World  ");
            var result = field.AsString();

            Assert.Equal("Hello World", result);
        }

        [Fact]
        public void Returns_Null_For_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsString();

            Assert.Null(result);
        }

        [Fact]
        public void Returns_Empty_String_For_Empty_Value()
        {
            var field = CreateField("");
            var result = field.AsString();

            Assert.Equal("", result);
        }
    }

    public class AsInteger : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_Numeric_String_To_Int()
        {
            var field = CreateField("123", XLDataType.Text);
            var result = field.AsInteger();

            Assert.Equal(123, result);
        }

        [Fact]
        public void Converts_Decimal_String_To_Int()
        {
            var field = CreateField("123,99", XLDataType.Text);
            var result = field.AsInteger();

            Assert.Equal(123, result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_String()
        {
            var field = CreateField("abc", XLDataType.Text);
            var result = field.AsInteger();

            Assert.Null(result);
        }

        [Fact]
        public void Returns_Null_For_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsInteger();

            Assert.Null(result);
        }
    }

    public class AsDecimal : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_Numeric_String_To_Decimal()
        {
            var field = CreateField("123.45", XLDataType.Text);
            var result = field.AsDecimal();

            Assert.Equal(123.45m, result);
        }

        [Fact]
        public void Converts_Russian_Number_Format()
        {
            var field = CreateField("123,45", XLDataType.Text);
            var result = field.AsDecimal();

            Assert.Equal(123.45m, result);
        }

        [Fact]
        public void Converts_Number_With_Thousands_Separator()
        {
            var field = CreateField("1 234,56", XLDataType.Text);
            var result = field.AsDecimal();

            Assert.Equal(1234.56m, result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_String()
        {
            var field = CreateField("abc", XLDataType.Text);
            var result = field.AsDecimal();

            Assert.Null(result);
        }

        [Fact]
        public void Returns_Null_For_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsDecimal();

            Assert.Null(result);
        }
    }

    public class AsDateTime : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_Date_String_To_DateTime()
        {
            var field = CreateField("15.01.2024", XLDataType.Text);
            var result = field.AsDateTime();

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Converts_Date_With_Slashes()
        {
            var field = CreateField("15/01/2024", XLDataType.Text);
            var result = field.AsDateTime();

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Converts_ISO_Date_Format()
        {
            var field = CreateField("2024-01-15", XLDataType.Text);
            var result = field.AsDateTime();

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_Date()
        {
            var field = CreateField("invalid", XLDataType.Text);
            var result = field.AsDateTime();

            Assert.Null(result);
        }

        [Fact]
        public void Returns_Null_For_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsDateTime();

            Assert.Null(result);
        }
    }

    public class AsBoolean : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_True_String_To_Bool()
        {
            var field = CreateField("true", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.True(result);
        }

        [Fact]
        public void Converts_False_String_To_Bool()
        {
            var field = CreateField("false", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.False(result);
        }

        [Fact]
        public void Converts_Russian_Da_To_True()
        {
            var field = CreateField("да", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.True(result);
        }

        [Fact]
        public void Converts_Russian_Net_To_False()
        {
            var field = CreateField("нет", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.False(result);
        }

        [Fact]
        public void Converts_1_To_True()
        {
            var field = CreateField("1", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.True(result);
        }

        [Fact]
        public void Converts_0_To_False()
        {
            var field = CreateField("0", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.False(result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_String()
        {
            var field = CreateField("invalid", XLDataType.Text);
            var result = field.AsBoolean();

            Assert.Null(result);
        }
    }

    public class AsDouble : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_Numeric_String_To_Double()
        {
            var field = CreateField("123.45", XLDataType.Text);
            var result = field.AsDouble();

            Assert.Equal(123.45, result);
        }

        [Fact]
        public void Converts_Russian_Number_Format()
        {
            var field = CreateField("123,45", XLDataType.Text);
            var result = field.AsDouble();

            Assert.Equal(123.45, result);
        }

        [Fact]
        public void Returns_Null_For_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsDouble();

            Assert.Null(result);
        }
    }

    public class AsLong : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_Large_Number_To_Long()
        {
            var field = CreateField("9223372036854775807", XLDataType.Number);
            var result = field.AsLong();

            Assert.Equal(long.MaxValue, result);
        }

        [Fact]
        public void Returns_Null_For_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsLong();

            Assert.Null(result);
        }
    }

    public class AsType : ParsedFieldExtensionsTests
    {
        [Fact]
        public void Converts_To_Specified_Type()
        {
            var field = CreateField("123,45", XLDataType.Text);
            var result = field.AsType<decimal>();

            Assert.Equal(123.45m, result);
        }

        [Fact]
        public void Converts_To_Int_Type()
        {
            var field = CreateField("123", XLDataType.Text);
            var result = field.AsType<int>();

            Assert.Equal(123, result);
        }

        [Fact]
        public void Converts_To_DateTime_Type()
        {
            var field = CreateField("15.01.2024", XLDataType.Text);
            var result = field.AsType<DateTime>();

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Returns_Null_For_Nullable_Type_With_Null_Value()
        {
            var field = CreateField(null);
            var result = field.AsType<int?>();

            Assert.Null(result);
        }
    }
}
