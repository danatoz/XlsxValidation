using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;

namespace XlsxValidation.Parsing;

/// <summary>
/// Парсер таблицы
/// </summary>
public class TableParser
{
    private readonly string _tableName;
    private readonly ICellAnchor _headerAnchor;
    private readonly StopConditionConfig? _stopCondition;
    private readonly int? _maxRows;
    private readonly List<ColumnParser> _columnParsers;
    private readonly TypeConverter _typeConverter;
    private int? _headerRowNumber;

    /// <summary>
    /// Создать парсер таблицы
    /// </summary>
    public TableParser(
        string tableName,
        ICellAnchor headerAnchor,
        StopConditionConfig? stopCondition,
        int? maxRows,
        List<ColumnParser> columnParsers,
        TypeConverter typeConverter)
    {
        _tableName = tableName;
        _headerAnchor = headerAnchor;
        _stopCondition = stopCondition;
        _maxRows = maxRows;
        _columnParsers = columnParsers;
        _typeConverter = typeConverter;
    }

    /// <summary>
    /// Распарсить таблицу из листа
    /// </summary>
    public ParsedTable Parse(IXLWorksheet worksheet, string worksheetName)
    {
        var errors = new List<ParseError>();

        // Найти строку заголовка
        var headerResult = _headerAnchor.Resolve(worksheet);
        if (!headerResult.IsSuccess || headerResult.Cell == null)
        {
            errors.Add(ParseError.Create(
                _tableName,
                $"Заголовок таблицы '{_tableName}' не найден",
                worksheetName: worksheetName));

            return new ParsedTable
            {
                Name = _tableName,
                Headers = new List<string>(),
                Rows = new List<ParsedTableRow>()
            };
        }

        _headerRowNumber = headerResult.Cell.Address.RowNumber;

        // Получить заголовки колонок
        var headers = ReadHeaders(worksheet, _headerRowNumber.Value);

        // Найти индексы колонок для парсеров
        foreach (var columnParser in _columnParsers)
        {
            columnParser.FindColumnIndex(headers, _headerRowNumber.Value);
        }

        // Читать строки данных
        var rows = ReadDataRows(worksheet, headers.Count);

        return new ParsedTable
        {
            Name = _tableName,
            Headers = headers,
            Rows = rows
        };
    }

    /// <summary>
    /// Прочитать заголовки колонок
    /// </summary>
    private List<string> ReadHeaders(IXLWorksheet worksheet, int headerRowNumber)
    {
        var headers = new List<string>();
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;

        for (int col = 1; col <= lastColumn; col++)
        {
            var cell = worksheet.Cell(headerRowNumber, col);
            var value = cell.GetValue<string>();

            if (string.IsNullOrWhiteSpace(value))
                break;

            headers.Add(value.Trim());
        }

        return headers;
    }

    /// <summary>
    /// Прочитать строки данных
    /// </summary>
    private List<ParsedTableRow> ReadDataRows(IXLWorksheet worksheet, int headerCount)
    {
        var rows = new List<ParsedTableRow>();

        if (_headerRowNumber == null)
            return rows;

        int currentRow = _headerRowNumber.Value + 1;
        int rowCount = 0;

        while (true)
        {
            // Проверка условия остановки
            if (ShouldStop(worksheet, currentRow, rowCount))
                break;

            // Проверка максимального количества строк
            if (_maxRows.HasValue && rowCount >= _maxRows.Value)
                break;

            // Пропустить пустую строку
            if (IsRowEmpty(worksheet, currentRow, headerCount))
            {
                if (_stopCondition?.Type == StopConditionType.EmptyRow)
                    break;

                currentRow++;
                continue;
            }

            // Распарсить строку
            var rowFields = new Dictionary<string, ParsedField>();

            foreach (var columnParser in _columnParsers)
            {
                var field = columnParser.Parse(worksheet, currentRow, _tableName);
                rowFields[field.Name] = field;
            }

            rows.Add(new ParsedTableRow
            {
                RowNumber = currentRow,
                Fields = rowFields
            });

            rowCount++;
            currentRow++;
        }

        return rows;
    }

    /// <summary>
    /// Проверить условие остановки
    /// </summary>
    private bool ShouldStop(IXLWorksheet worksheet, int rowNumber, int rowCount)
    {
        if (_stopCondition == null)
            return false;

        return _stopCondition.Type switch
        {
            StopConditionType.EmptyRow => IsRowEmpty(worksheet, rowNumber, 1),
            StopConditionType.SentinelValue => HasSentinelValue(worksheet, rowNumber),
            StopConditionType.MaxRows => _maxRows.HasValue && rowCount >= _maxRows.Value,
            _ => false
        };
    }

    /// <summary>
    /// Проверить наличие значения-маркера
    /// </summary>
    private bool HasSentinelValue(IXLWorksheet worksheet, int rowNumber)
    {
        if (_stopCondition?.SentinelValue == null)
            return false;

        var firstCell = worksheet.Cell(rowNumber, 1);
        var value = firstCell.GetValue<string>()?.Trim();

        return value == _stopCondition.SentinelValue;
    }

    /// <summary>
    /// Проверить, пуста ли строка
    /// </summary>
    private bool IsRowEmpty(IXLWorksheet worksheet, int rowNumber, int columnsToCheck)
    {
        for (int col = 1; col <= columnsToCheck; col++)
        {
            var cell = worksheet.Cell(rowNumber, col);
            if (!cell.IsEmpty())
                return false;
        }

        return true;
    }
}
