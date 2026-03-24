using System.Globalization;
using ClosedXML.Excel;
using XlsxValidation.Configuration;

namespace XlsxValidation.Parsing;

/// <summary>
/// Конвертер типов данных для парсинга XLSX
/// </summary>
public class TypeConverter
{
    private readonly CultureInfo _culture;
    private readonly string[] _dateFormats;
    private readonly NumberStyles _numberStyles;
    private readonly bool _trimStrings;

    /// <summary>
    /// Создать конвертер с опциями
    /// </summary>
    public TypeConverter(ParseOptions options)
    {
        _culture = CultureInfo.GetCultureInfo(options.Culture);
        _dateFormats = options.DateFormats;
        _numberStyles = options.NumberStyles;
        _trimStrings = options.TrimStrings;
    }

    /// <summary>
    /// Конвертировать значение в указанный тип
    /// </summary>
    public object? Convert(string? value, XLDataType dataType, Type targetType)
    {
        if (string.IsNullOrWhiteSpace(value) && dataType != XLDataType.Text)
            return GetDefault(targetType);

        var trimmedValue = _trimStrings ? value?.Trim() : value;

        if (targetType == typeof(string))
            return ConvertToString(trimmedValue);

        if (targetType == typeof(int) || targetType == typeof(int?))
            return ConvertToInt(trimmedValue, dataType);

        if (targetType == typeof(long) || targetType == typeof(long?))
            return ConvertToLong(trimmedValue, dataType);

        if (targetType == typeof(decimal) || targetType == typeof(decimal?))
            return ConvertToDecimal(trimmedValue, dataType);

        if (targetType == typeof(double) || targetType == typeof(double?))
            return ConvertToDouble(trimmedValue, dataType);

        if (targetType == typeof(DateTime) || targetType == typeof(DateTime?))
            return ConvertToDateTime(trimmedValue, dataType);

        if (targetType == typeof(DateOnly) || targetType == typeof(DateOnly?))
            return ConvertToDateOnly(trimmedValue, dataType);

        if (targetType == typeof(bool) || targetType == typeof(bool?))
            return ConvertToBoolean(trimmedValue, dataType);

        if (targetType == typeof(TimeSpan) || targetType == typeof(TimeSpan?))
            return ConvertToTimeSpan(trimmedValue, dataType);

        // Попытка конвертации через ChangeType
        try
        {
            if (string.IsNullOrWhiteSpace(trimmedValue))
                return GetDefault(targetType);

            return System.Convert.ChangeType(trimmedValue, targetType, _culture);
        }
        catch
        {
            return GetDefault(targetType);
        }
    }

    /// <summary>
    /// Конвертировать в строку
    /// </summary>
    public string? ConvertToString(string? value)
    {
        if (value == null)
            return null;

        return _trimStrings ? value.Trim() : value;
    }

    /// <summary>
    /// Конвертировать в decimal
    /// </summary>
    public decimal? ConvertToDecimal(string? value, XLDataType dataType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var normalizedValue = NormalizeNumberString(value);

        // Парсим с инвариантной культурой (точка как разделитель)
        if (decimal.TryParse(normalizedValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
            return result;

        return null;
    }

    /// <summary>
    /// Конвертировать в int
    /// </summary>
    public int? ConvertToInt(string? value, XLDataType dataType)
    {
        var decimalValue = ConvertToDecimal(value, dataType);
        if (decimalValue.HasValue)
            return (int)decimalValue.Value;

        return null;
    }

    /// <summary>
    /// Конвертировать в long
    /// </summary>
    public long? ConvertToLong(string? value, XLDataType dataType)
    {
        var decimalValue = ConvertToDecimal(value, dataType);
        if (decimalValue.HasValue)
            return (long)decimalValue.Value;

        return null;
    }

    /// <summary>
    /// Конвертировать в double
    /// </summary>
    public double? ConvertToDouble(string? value, XLDataType dataType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var normalizedValue = NormalizeNumberString(value);

        // Парсим с инвариантной культурой (точка как разделитель)
        if (double.TryParse(normalizedValue, NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
            return result;

        return null;
    }

    /// <summary>
    /// Конвертировать в DateTime
    /// </summary>
    public DateTime? ConvertToDateTime(string? value, XLDataType dataType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        if (dataType == XLDataType.DateTime)
        {
            if (DateTime.TryParse(value, _culture, DateTimeStyles.None, out var dtResult))
                return dtResult;
        }

        if (dataType == XLDataType.Text || dataType == XLDataType.Number)
        {
            // Попытка парсинга по заданным форматам
            foreach (var format in _dateFormats)
            {
                if (DateTime.TryParseExact(value, format, _culture, DateTimeStyles.None, out var formatResult))
                    return formatResult;
            }

            // Попытка общего парсинга
            if (DateTime.TryParse(value, _culture, DateTimeStyles.None, out var parseResult))
                return parseResult;
        }

        return null;
    }

    /// <summary>
    /// Конвертировать в DateOnly
    /// </summary>
    public DateOnly? ConvertToDateOnly(string? value, XLDataType dataType)
    {
        var dateTime = ConvertToDateTime(value, dataType);
        if (dateTime.HasValue)
            return DateOnly.FromDateTime(dateTime.Value);

        return null;
    }

    /// <summary>
    /// Конвертировать в bool
    /// </summary>
    public bool? ConvertToBoolean(string? value, XLDataType dataType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        var trimmed = value.Trim().ToLowerInvariant();

        // Проверка русских значений
        if (trimmed == "да" || trimmed == "истина")
            return true;
        if (trimmed == "нет" || trimmed == "ложь")
            return false;

        // Стандартные значения
        if (bool.TryParse(trimmed, out var result))
            return result;

        // Числовые значения
        if (int.TryParse(trimmed, out var num))
            return num != 0;

        return null;
    }

    /// <summary>
    /// Конвертировать в TimeSpan
    /// </summary>
    public TimeSpan? ConvertToTimeSpan(string? value, XLDataType dataType)
    {
        if (string.IsNullOrWhiteSpace(value))
            return null;

        if (TimeSpan.TryParse(value, _culture, out var result))
            return result;

        return null;
    }

    /// <summary>
    /// Получить значение по умолчанию для типа
    /// </summary>
    private object? GetDefault(Type type)
    {
        if (type.IsValueType && Nullable.GetUnderlyingType(type) == null)
            return Activator.CreateInstance(type);

        return null;
    }

    /// <summary>
    /// Нормализовать строку числа (замена разделителей)
    /// </summary>
    private string NormalizeNumberString(string value)
    {
        // Заменяем возможные разделители тысяч и десятичных
        var normalized = value.Replace(" ", "")
            .Replace("\u00A0", "") // неразрывный пробел
            .Trim();

        // Определяем формат числа и нормализуем к культуре
        // Если есть и точка и запятая, то запятая - это разделитель тысяч
        if (normalized.Contains(",") && normalized.Contains("."))
        {
            // 1,234.56 или 1.234,56
            var lastSeparatorIndex = Math.Max(normalized.LastIndexOf(','), normalized.LastIndexOf('.'));
            var beforeLast = normalized.Substring(0, lastSeparatorIndex);
            var afterLast = normalized.Substring(lastSeparatorIndex + 1);

            // Если после последнего разделителя 3 цифры или меньше, это десятичная часть
            if (afterLast.Length <= 3 && afterLast.All(char.IsDigit))
            {
                // Это десятичный разделитель, удаляем разделители тысяч
                beforeLast = beforeLast.Replace(",", "").Replace(".", "");
                normalized = beforeLast + "." + afterLast;
            }
            else
            {
                // Оба являются разделителями тысяч
                normalized = beforeLast.Replace(",", "").Replace(".", "") + "." + afterLast;
            }
        }
        else if (normalized.Contains(","))
        {
            // Только запятая - заменяем на точку для универсального парсинга
            normalized = normalized.Replace(",", ".");
        }

        return normalized;
    }
}
