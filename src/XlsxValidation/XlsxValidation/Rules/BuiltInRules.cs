using System.Globalization;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using XlsxValidation.Configuration;
using XlsxValidation.Results;

namespace XlsxValidation.Rules;

/// <summary>
/// Встроенные правила валидации
/// </summary>
public static class BuiltInRules
{
    /// <summary>
    /// Зарегистрировать все встроенные правила в реестре
    /// </summary>
    public static void RegisterDefaults(XlsxRuleRegistry registry)
    {
        // === Shared Rules (для ячеек и колонок) ===

        // not-empty
        registry.RegisterSharedRule("not-empty", (cfg, prefix) => cell =>
        {
            var value = cell.GetValue<string>();
            if (string.IsNullOrWhiteSpace(value))
                return ValidationResult.Error($"{prefix}Ячейка не должна быть пустой");
            return ValidationResult.Ok();
        });

        // is-numeric
        registry.RegisterSharedRule("is-numeric", (cfg, prefix) => cell =>
        {
            if (cell.DataType == XLDataType.Number)
                return ValidationResult.Ok();
            
            var value = cell.GetValue<string>();
            if (string.IsNullOrWhiteSpace(value))
                return ValidationResult.Ok(); // Пустая ячейка проверяется not-empty
            
            if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out _) ||
                double.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out _))
                return ValidationResult.Ok();
            
            return ValidationResult.Error($"{prefix}Значение должно быть числом");
        });

        // is-date
        registry.RegisterSharedRule("is-date", (cfg, prefix) => cell =>
        {
            if (cell.DataType == XLDataType.DateTime)
                return ValidationResult.Ok();
            
            var value = cell.GetValue<string>();
            if (string.IsNullOrWhiteSpace(value))
                return ValidationResult.Ok();
            
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                return ValidationResult.Ok();
            
            return ValidationResult.Error($"{prefix}Значение должно быть датой");
        });

        // is-text
        registry.RegisterSharedRule("is-text", (cfg, prefix) => cell =>
        {
            var value = cell.GetValue<string>();
            if (string.IsNullOrEmpty(value))
                return ValidationResult.Ok();
            
            // Если это число или дата, то это не текст
            if (cell.DataType == XLDataType.Number || cell.DataType == XLDataType.DateTime)
                return ValidationResult.Error($"{prefix}Значение должно быть строкой");
            
            return ValidationResult.Ok();
        });

        // max-length
        registry.RegisterSharedRule("max-length", (cfg, prefix) => cell =>
        {
            var value = cell.GetValue<string>();
            if (string.IsNullOrEmpty(value))
                return ValidationResult.Ok();

            if (!cfg.Params.TryGetValue("max", out var maxObj) || maxObj is not int max)
                return ValidationResult.Error($"{prefix}Правило max-length требует параметр 'max'");

            if (value.Length > max)
                return ValidationResult.Error($"{prefix}Длина не должна превышать {max} символов");
            
            return ValidationResult.Ok();
        });

        // min-length
        registry.RegisterSharedRule("min-length", (cfg, prefix) => cell =>
        {
            var value = cell.GetValue<string>();
            if (string.IsNullOrEmpty(value))
                return ValidationResult.Ok();

            if (!cfg.Params.TryGetValue("min", out var minObj) || minObj is not int min)
                return ValidationResult.Error($"{prefix}Правило min-length требует параметр 'min'");

            if (value.Length < min)
                return ValidationResult.Error($"{prefix}Длина должна быть не менее {min} символов");
            
            return ValidationResult.Ok();
        });

        // min-value
        registry.RegisterSharedRule("min-value", (cfg, prefix) => cell =>
        {
            if (!cfg.Params.TryGetValue("min", out var minObj))
                return ValidationResult.Error($"{prefix}Правило min-value требует параметр 'min'");

            double? value = GetNumericValue(cell);
            if (value == null)
                return ValidationResult.Ok(); // Пустая или нечисловая ячейка

            double min = minObj switch
            {
                int i => i,
                double d => d,
                decimal m => (double)m,
                _ => throw new ArgumentException($"Неверный тип параметра min: {minObj.GetType()}")
            };

            if (value < min)
                return ValidationResult.Error($"{prefix}Значение должно быть >= {min}");
            
            return ValidationResult.Ok();
        });

        // max-value
        registry.RegisterSharedRule("max-value", (cfg, prefix) => cell =>
        {
            if (!cfg.Params.TryGetValue("max", out var maxObj))
                return ValidationResult.Error($"{prefix}Правило max-value требует параметр 'max'");

            double? value = GetNumericValue(cell);
            if (value == null)
                return ValidationResult.Ok();

            double max = maxObj switch
            {
                int i => i,
                double d => d,
                decimal m => (double)m,
                _ => throw new ArgumentException($"Неверный тип параметра max: {maxObj.GetType()}")
            };

            if (value > max)
                return ValidationResult.Error($"{prefix}Значение должно быть <= {max}");
            
            return ValidationResult.Ok();
        });

        // matches (regex)
        registry.RegisterSharedRule("matches", (cfg, prefix) => cell =>
        {
            var value = cell.GetValue<string>();
            if (string.IsNullOrEmpty(value))
                return ValidationResult.Ok();

            if (!cfg.Params.TryGetValue("pattern", out var patternObj) || patternObj is not string pattern)
                return ValidationResult.Error($"{prefix}Правило matches требует параметр 'pattern'");

            var message = cfg.Params.TryGetValue("message", out var msgObj) && msgObj is string msg
                ? $"{prefix}{msg}"
                : $"{prefix}Значение не соответствует шаблону '{pattern}'";

            if (!Regex.IsMatch(value, pattern))
                return ValidationResult.Error(message);
            
            return ValidationResult.Ok();
        });

        // one-of
        registry.RegisterSharedRule("one-of", (cfg, prefix) => cell =>
        {
            var value = cell.GetValue<string>();
            if (string.IsNullOrEmpty(value))
                return ValidationResult.Ok();

            if (!cfg.Params.TryGetValue("values", out var valuesObj) || valuesObj is not IEnumerable<object> values)
                return ValidationResult.Error($"{prefix}Правило one-of требует параметр 'values'");

            var allowedValues = values.Select(v => v?.ToString()?.ToLowerInvariant()).ToHashSet();
            if (!allowedValues.Contains(value.Trim().ToLowerInvariant()))
                return ValidationResult.Error($"{prefix}Значение должно быть одним из: {string.Join(", ", values.Select(v => $"'{v}'"))}");
            
            return ValidationResult.Ok();
        });

        // === Cell-Only Rules ===

        // date-not-future
        registry.RegisterCellRule("date-not-future", cfg => cell =>
        {
            DateTime? date = GetDateValue(cell);
            if (date == null)
                return ValidationResult.Ok();

            if (date.Value.Date > DateTime.Today)
                return ValidationResult.Error("Дата не может быть в будущем");
            
            return ValidationResult.Ok();
        });

        // date-not-past
        registry.RegisterCellRule("date-not-past", cfg => cell =>
        {
            DateTime? date = GetDateValue(cell);
            if (date == null)
                return ValidationResult.Ok();

            if (date.Value.Date < DateTime.Today)
                return ValidationResult.Error("Дата не может быть в прошлом");
            
            return ValidationResult.Ok();
        });

        // is-merged
        registry.RegisterCellRule("is-merged", cfg => cell =>
        {
            var range = cell.MergedRange;
            if (range == null)
                return ValidationResult.Error("Ячейка должна быть объединённой");
            
            return ValidationResult.Ok();
        });
    }

    /// <summary>
    /// Получить числовое значение из ячейки
    /// </summary>
    private static double? GetNumericValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return null;

        if (cell.DataType == XLDataType.Number)
            return cell.GetValue<double>();

        var value = cell.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value))
            return null;

        if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out var result))
            return result;
        
        if (double.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out result))
            return result;

        return null;
    }

    /// <summary>
    /// Получить значение даты из ячейки
    /// </summary>
    private static DateTime? GetDateValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return null;

        if (cell.DataType == XLDataType.DateTime)
            return cell.GetValue<DateTime>();

        var value = cell.GetValue<string>();
        if (string.IsNullOrWhiteSpace(value))
            return null;

        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out var result))
            return result;

        return null;
    }
}
