namespace XlsxValidation.Configuration;

/// <summary>
/// Опции парсинга XLSX файлов
/// </summary>
public record ParseOptions
{
    /// <summary>
    /// Пропускать ли пустые ячейки
    /// </summary>
    public bool SkipEmptyCells { get; init; } = true;

    /// <summary>
    /// Обрезать ли пробелы у строковых значений
    /// </summary>
    public bool TrimStrings { get; init; } = true;

    /// <summary>
    /// Культура для парсинга чисел и дат (например, "ru-RU")
    /// </summary>
    public string Culture { get; init; } = "ru-RU";

    /// <summary>
    /// Форматы дат для парсинга
    /// </summary>
    public string[] DateFormats { get; init; } = new[] { "dd.MM.yyyy", "dd/MM/yyyy", "yyyy-MM-dd" };

    /// <summary>
    /// Стиль парсинга чисел
    /// </summary>
    public System.Globalization.NumberStyles NumberStyles { get; init; } = System.Globalization.NumberStyles.Number;

    /// <summary>
    /// Значения по умолчанию для типов
    /// </summary>
    public Dictionary<string, object> Defaults { get; init; } = new();
}

/// <summary>
/// Маппинг поля на тип данных
/// </summary>
public record FieldMappingConfig
{
    /// <summary>
    /// Имя поля
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Тип данных (string, int, decimal, date, bool, etc.)
    /// </summary>
    public string Type { get; init; } = "string";

    /// <summary>
    /// Имя колонки (для табличных полей)
    /// </summary>
    public string? Column { get; init; }
}

/// <summary>
/// Маппинг колонки таблицы
/// </summary>
public record ColumnMappingConfig
{
    /// <summary>
    /// Имя свойства в модели
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Тип данных
    /// </summary>
    public string Type { get; init; } = "string";
}

/// <summary>
/// Маппинг таблицы
/// </summary>
public record TableMappingConfig
{
    /// <summary>
    /// Имя таблицы
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Тип строки (для маппинга на модель)
    /// </summary>
    public string RowType { get; init; } = string.Empty;

    /// <summary>
    /// Маппинг колонок
    /// </summary>
    public Dictionary<string, ColumnMappingConfig> Columns { get; init; } = new();
}

/// <summary>
/// Секция парсинга в профиле валидации
/// </summary>
public record ParsingSection
{
    /// <summary>
    /// Маппинг полей на типы данных
    /// </summary>
    public Dictionary<string, string> FieldTypes { get; init; } = new();

    /// <summary>
    /// Опции парсинга
    /// </summary>
    public ParseOptions Options { get; init; } = new();

    /// <summary>
    /// Маппинг таблиц
    /// </summary>
    public Dictionary<string, TableMappingConfig> TableMapping { get; init; } = new();
}
