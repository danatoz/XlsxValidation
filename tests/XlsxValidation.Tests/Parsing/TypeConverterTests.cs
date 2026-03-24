using System.Globalization;
using ClosedXML.Excel;
using XlsxValidation.Configuration;
using XlsxValidation.Parsing;

namespace XlsxValidation.Tests.Parsing;

public class TypeConverterTests
{
    private static TypeConverter CreateConverter()
    {
        return new TypeConverter(new ParseOptions
        {
            Culture = "ru-RU",
            TrimStrings = true,
            DateFormats = new[] { "dd.MM.yyyy", "dd/MM/yyyy", "yyyy-MM-dd" },
            NumberStyles = NumberStyles.Number
        });
    }

    public class ConvertToString : TypeConverterTests
    {
        [Fact]
        public void Returns_String_As_Is()
        {
            var converter = CreateConverter();
            var result = converter.ConvertToString("  Hello World  ");

            Assert.Equal("Hello World", result);
        }

        [Fact]
        public void Returns_Null_For_Null()
        {
            var converter = CreateConverter();
            var result = converter.ConvertToString(null);

            Assert.Null(result);
        }
    }

    public class ConvertToDecimal : TypeConverterTests
    {
        [Fact]
        public void Converts_Number_String_To_Decimal()
        {
            var converter = CreateConverter();
            // Для XLDataType.Number значение уже числовое, передаем как строку с точкой
            var result = converter.ToDecimal("123.45", XLDataType.Text);

            Assert.Equal(123.45m, result);
        }

        [Fact]
        public void Converts_Russian_Number_String_To_Decimal()
        {
            var converter = CreateConverter();
            var result = converter.ToDecimal("123,45", XLDataType.Text);

            Assert.Equal(123.45m, result);
        }

        [Fact]
        public void Returns_Null_For_Empty_String()
        {
            var converter = CreateConverter();
            var result = converter.ToDecimal("", XLDataType.Text);

            Assert.Null(result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_String()
        {
            var converter = CreateConverter();
            var result = converter.ToDecimal("abc", XLDataType.Text);

            Assert.Null(result);
        }

        [Fact]
        public void Converts_Number_With_Thousands_Separator()
        {
            var converter = CreateConverter();
            var result = converter.ToDecimal("1 234,56", XLDataType.Text);

            Assert.Equal(1234.56m, result);
        }
    }

    public class ConvertToInt : TypeConverterTests
    {
        [Fact]
        public void Converts_Number_String_To_Int()
        {
            var converter = CreateConverter();
            var result = converter.ToInt("123", XLDataType.Text);

            Assert.Equal(123, result);
        }

        [Fact]
        public void Converts_Decimal_String_To_Int()
        {
            var converter = CreateConverter();
            var result = converter.ToInt("123,99", XLDataType.Text);

            Assert.Equal(123, result);
        }

        [Fact]
        public void Returns_Null_For_Empty_String()
        {
            var converter = CreateConverter();
            var result = converter.ToInt("", XLDataType.Text);

            Assert.Null(result);
        }
    }

    public class ConvertToDateTime : TypeConverterTests
    {
        [Fact]
        public void Converts_Date_String_To_DateTime()
        {
            var converter = CreateConverter();
            var result = converter.ToDateTime("15.01.2024", XLDataType.Text);

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Converts_Date_With_Slashes()
        {
            var converter = CreateConverter();
            var result = converter.ToDateTime("15/01/2024", XLDataType.Text);

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Converts_ISO_Date_Format()
        {
            var converter = CreateConverter();
            var result = converter.ToDateTime("2024-01-15", XLDataType.Text);

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_Date()
        {
            var converter = CreateConverter();
            var result = converter.ToDateTime("invalid", XLDataType.Text);

            Assert.Null(result);
        }

        [Fact]
        public void Returns_Null_For_Empty_String()
        {
            var converter = CreateConverter();
            var result = converter.ToDateTime("", XLDataType.Text);

            Assert.Null(result);
        }
    }

    public class ConvertToBoolean : TypeConverterTests
    {
        [Fact]
        public void Converts_True_String_To_Bool()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("true", XLDataType.Text);

            Assert.True(result);
        }

        [Fact]
        public void Converts_False_String_To_Bool()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("false", XLDataType.Text);

            Assert.False(result);
        }

        [Fact]
        public void Converts_Russian_Da_To_True()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("да", XLDataType.Text);

            Assert.True(result);
        }

        [Fact]
        public void Converts_Russian_Net_To_False()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("нет", XLDataType.Text);

            Assert.False(result);
        }

        [Fact]
        public void Converts_1_To_True()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("1", XLDataType.Text);

            Assert.True(result);
        }

        [Fact]
        public void Converts_0_To_False()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("0", XLDataType.Text);

            Assert.False(result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_String()
        {
            var converter = CreateConverter();
            var result = converter.ToBoolean("invalid", XLDataType.Text);

            Assert.Null(result);
        }
    }

    public class ConvertToDouble : TypeConverterTests
    {
        [Fact]
        public void Converts_Number_String_To_Double()
        {
            var converter = CreateConverter();
            var result = converter.ToDouble("123.45", XLDataType.Text);

            Assert.Equal(123.45, result);
        }

        [Fact]
        public void Converts_Russian_Number_String_To_Double()
        {
            var converter = CreateConverter();
            var result = converter.ToDouble("123,45", XLDataType.Text);

            Assert.Equal(123.45, result);
        }

        [Fact]
        public void Returns_Null_For_Empty_String()
        {
            var converter = CreateConverter();
            var result = converter.ToDouble("", XLDataType.Text);

            Assert.Null(result);
        }
    }

    public class ConvertToLong : TypeConverterTests
    {
        [Fact]
        public void Converts_Number_String_To_Long()
        {
            var converter = CreateConverter();
            var result = converter.ToLong("9223372036854775807", XLDataType.Text);

            Assert.Equal(long.MaxValue, result);
        }

        [Fact]
        public void Converts_Decimal_String_To_Long()
        {
            var converter = CreateConverter();
            var result = converter.ToLong("123,99", XLDataType.Text);

            Assert.Equal(123L, result);
        }
    }

    public class ConvertToTimeSpan : TypeConverterTests
    {
        [Fact]
        public void Converts_Time_String_To_TimeSpan()
        {
            var converter = CreateConverter();
            var result = converter.ToTimeSpan("01:30:00", XLDataType.Text);

            Assert.Equal(TimeSpan.FromHours(1.5), result);
        }

        [Fact]
        public void Returns_Null_For_Invalid_Time()
        {
            var converter = CreateConverter();
            var result = converter.ToTimeSpan("invalid", XLDataType.Text);

            Assert.Null(result);
        }
    }

    public class ConvertToGeneric : TypeConverterTests
    {
        [Fact]
        public void Converts_To_Specified_Type()
        {
            var converter = CreateConverter();
            var result = converter.Convert("123,45", XLDataType.Text, typeof(decimal));

            Assert.Equal(123.45m, result);
        }

        [Fact]
        public void Converts_To_Int_Type()
        {
            var converter = CreateConverter();
            var result = converter.Convert("123", XLDataType.Text, typeof(int));

            Assert.Equal(123, result);
        }

        [Fact]
        public void Converts_To_DateTime_Type()
        {
            var converter = CreateConverter();
            var result = converter.Convert("15.01.2024", XLDataType.Text, typeof(DateTime));

            Assert.Equal(new DateTime(2024, 1, 15), result);
        }

        [Fact]
        public void Returns_Default_For_Nullable_Type()
        {
            var converter = CreateConverter();
            var result = converter.Convert("", XLDataType.Text, typeof(int?));

            Assert.Null(result);
        }
    }
}

/// <summary>
/// Extension-методы для тестирования TypeConverter
/// </summary>
file static class TypeConverterExtensions
{
    public static decimal? ToDecimal(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToDecimal(value, dataType);

    public static int? ToInt(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToInt(value, dataType);

    public static long? ToLong(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToLong(value, dataType);

    public static double? ToDouble(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToDouble(value, dataType);

    public static DateTime? ToDateTime(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToDateTime(value, dataType);

    public static bool? ToBoolean(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToBoolean(value, dataType);

    public static TimeSpan? ToTimeSpan(this TypeConverter converter, string value, XLDataType dataType)
        => converter.ConvertToTimeSpan(value, dataType);
}
