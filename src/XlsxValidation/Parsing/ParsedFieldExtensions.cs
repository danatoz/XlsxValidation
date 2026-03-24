using System.Globalization;
using ClosedXML.Excel;

namespace XlsxValidation.Parsing;

/// <summary>
/// Extension-методы для ParsedField
/// </summary>
public static class ParsedFieldExtensions
{
    /// <summary>
    /// Получить значение как строку
    /// </summary>
    public static string? AsString(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        return field.RawValue.Trim();
    }

    /// <summary>
    /// Получить значение как integer
    /// </summary>
    public static int? AsInteger(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToInt(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как long
    /// </summary>
    public static long? AsLong(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToLong(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как decimal
    /// </summary>
    public static decimal? AsDecimal(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToDecimal(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как double
    /// </summary>
    public static double? AsDouble(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToDouble(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как DateTime
    /// </summary>
    public static DateTime? AsDateTime(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToDateTime(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как DateOnly
    /// </summary>
    public static DateOnly? AsDateOnly(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToDateOnly(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как bool
    /// </summary>
    public static bool? AsBoolean(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToBoolean(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как TimeSpan
    /// </summary>
    public static TimeSpan? AsTimeSpan(this ParsedField field)
    {
        if (field.RawValue == null)
            return null;

        var converter = CreateDefaultConverter();
        return converter.ConvertToTimeSpan(field.RawValue, field.DataType);
    }

    /// <summary>
    /// Получить значение как указанный тип
    /// </summary>
    public static T? AsType<T>(this ParsedField field)
    {
        if (field.RawValue == null)
            return default;

        var converter = CreateDefaultConverter();
        var targetType = typeof(T);
        var converted = converter.Convert(field.RawValue, field.DataType, targetType);

        if (converted == null)
            return default;

        return (T)converted;
    }

    /// <summary>
    /// Получить значение как указанный тип (для не nullable типов возвращает default)
    /// </summary>
    public static T? AsType<T>(this ParsedField field, TypeConverter converter)
    {
        if (field.RawValue == null)
            return default;

        var targetType = typeof(T);
        var converted = converter.Convert(field.RawValue, field.DataType, targetType);

        if (converted == null)
            return default;

        return (T)converted;
    }

    /// <summary>
    /// Конвертировать поле в указанный тип
    /// </summary>
    public static object? AsType(this ParsedField field, Type targetType)
    {
        return AsType(field, targetType, null);
    }

    /// <summary>
    /// Конвертировать поле в указанный тип с использованием конвертера
    /// </summary>
    public static object? AsType(this ParsedField field, Type targetType, TypeConverter? converter)
    {
        if (field.RawValue == null)
            return null;

        var actualConverter = converter ?? CreateDefaultConverter();
        return actualConverter.Convert(field.RawValue, field.DataType, targetType);
    }

    /// <summary>
    /// Создать конвертер по умолчанию
    /// </summary>
    private static TypeConverter CreateDefaultConverter()
    {
        return new TypeConverter(new Configuration.ParseOptions
        {
            Culture = "ru-RU",
            TrimStrings = true,
            DateFormats = new[] { "dd.MM.yyyy", "dd/MM/yyyy", "yyyy-MM-dd" },
            NumberStyles = NumberStyles.Number
        });
    }
}
