using ClosedXML.Excel;

namespace XlsxValidation.Parsing;

/// <summary>
/// Распаршенное поле (одиночная ячейка)
/// </summary>
public record ParsedField
{
    /// <summary>
    /// Логическое имя поля
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Сырое значение из ячейки
    /// </summary>
    public string? RawValue { get; init; }

    /// <summary>
    /// Тип данных ClosedXML
    /// </summary>
    public XLDataType DataType { get; init; }

    /// <summary>
    /// Адрес ячейки (например, "B3")
    /// </summary>
    public string? CellAddress { get; init; }

    /// <summary>
    /// Имя листа
    /// </summary>
    public string? WorksheetName { get; init; }

    /// <summary>
    /// Пустая ячейка или нет
    /// </summary>
    public bool IsEmpty => string.IsNullOrWhiteSpace(RawValue);
}

/// <summary>
/// Строка таблицы
/// </summary>
public record ParsedTableRow
{
    /// <summary>
    /// Номер строки в листе
    /// </summary>
    public int RowNumber { get; init; }

    /// <summary>
    /// Поля строки (ключ - заголовок колонки)
    /// </summary>
    public Dictionary<string, ParsedField> Fields { get; init; } = new();
}

/// <summary>
/// Распаршенная таблица
/// </summary>
public record ParsedTable
{
    /// <summary>
    /// Логическое имя таблицы
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Заголовки колонок
    /// </summary>
    public IReadOnlyList<string> Headers { get; init; } = new List<string>();

    /// <summary>
    /// Строки таблицы
    /// </summary>
    public IReadOnlyList<ParsedTableRow> Rows { get; init; } = new List<ParsedTableRow>();

    /// <summary>
    /// Количество строк
    /// </summary>
    public int RowCount => Rows.Count;

    /// <summary>
    /// Получить строку по индексу
    /// </summary>
    public ParsedTableRow? GetRow(int index) => index >= 0 && index < Rows.Count ? Rows[index] : null;
}

/// <summary>
/// Ошибка парсинга
/// </summary>
public record ParseError
{
    /// <summary>
    /// Имя поля, где произошла ошибка
    /// </summary>
    public string FieldName { get; init; } = string.Empty;

    /// <summary>
    /// Адрес ячейки (если применимо)
    /// </summary>
    public string? CellAddress { get; init; }

    /// <summary>
    /// Номер строки (если применимо)
    /// </summary>
    public int? RowNumber { get; init; }

    /// <summary>
    /// Сообщение об ошибке
    /// </summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>
    /// Исключение (если есть)
    /// </summary>
    public Exception? Exception { get; init; }

    /// <summary>
    /// Имя листа
    /// </summary>
    public string? WorksheetName { get; init; }

    /// <summary>
    /// Создать успешный результат
    /// </summary>
    public static ParseError Create(string fieldName, string message, string? cellAddress = null, int? rowNumber = null, string? worksheetName = null, Exception? exception = null)
        => new()
        {
            FieldName = fieldName,
            Message = message,
            CellAddress = cellAddress,
            RowNumber = rowNumber,
            WorksheetName = worksheetName,
            Exception = exception
        };
}

/// <summary>
/// Результат парсинга XLSX файла
/// </summary>
public record XlsxParseResult
{
    /// <summary>
    /// Имя профиля валидации
    /// </summary>
    public string ProfileName { get; init; } = string.Empty;

    /// <summary>
    /// Распаршенные одиночные поля
    /// </summary>
    public IReadOnlyList<ParsedField> Fields { get; init; } = new List<ParsedField>();

    /// <summary>
    /// Распаршенные таблицы
    /// </summary>
    public IReadOnlyList<ParsedTable> Tables { get; init; } = new List<ParsedTable>();

    /// <summary>
    /// Ошибки парсинга
    /// </summary>
    public IReadOnlyList<ParseError> Errors { get; init; } = new List<ParseError>();

    /// <summary>
    /// Был ли парсинг успешным
    /// </summary>
    public bool IsSuccess => Errors.Count == 0;

    /// <summary>
    /// Создать пустой результат
    /// </summary>
    public static XlsxParseResult Empty(string profileName) => new()
    {
        ProfileName = profileName,
        Fields = new List<ParsedField>(),
        Tables = new List<ParsedTable>(),
        Errors = new List<ParseError>()
    };

    /// <summary>
    /// Создать результат с ошибками
    /// </summary>
    public static XlsxParseResult WithErrors(string profileName, IEnumerable<ParseError> errors) => new()
    {
        ProfileName = profileName,
        Fields = new List<ParsedField>(),
        Tables = new List<ParsedTable>(),
        Errors = errors.ToList().AsReadOnly()
    };

    /// <summary>
    /// Создать успешный результат
    /// </summary>
    public static XlsxParseResult Success(
        string profileName,
        IEnumerable<ParsedField> fields,
        IEnumerable<ParsedTable> tables) => new()
    {
        ProfileName = profileName,
        Fields = fields.ToList().AsReadOnly(),
        Tables = tables.ToList().AsReadOnly(),
        Errors = new List<ParseError>()
    };

    /// <summary>
    /// Получить поле по имени
    /// </summary>
    public ParsedField? GetField(string name) => Fields.FirstOrDefault(f => f.Name == name);

    /// <summary>
    /// Получить таблицу по имени
    /// </summary>
    public ParsedTable? GetTable(string name) => Tables.FirstOrDefault(t => t.Name == name);
}
