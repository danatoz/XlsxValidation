using ClosedXML.Excel;
using XlsxValidation.Configuration;

namespace XlsxValidation.Parsing;

/// <summary>
/// Парсер колонки таблицы
/// </summary>
public class ColumnParser
{
    private readonly string _header;
    private readonly string _fieldType;
    private readonly TypeConverter _typeConverter;
    private int? _columnIndex;

    /// <summary>
    /// Создать парсер колонки
    /// </summary>
    public ColumnParser(
        string header,
        string fieldType,
        TypeConverter typeConverter)
    {
        _header = header;
        _fieldType = fieldType;
        _typeConverter = typeConverter;
    }

    /// <summary>
    /// Найти индекс колонки по заголовку
    /// </summary>
    public void FindColumnIndex(IEnumerable<string> headers, int headerRowNumber)
    {
        var headersList = headers.ToList();
        _columnIndex = headersList.FindIndex(h =>
            h.Equals(_header, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// Распарсить значение ячейки в строке
    /// </summary>
    public ParsedField Parse(IXLWorksheet worksheet, int rowNumber, string tableName)
    {
        if (!_columnIndex.HasValue || _columnIndex.Value < 0)
        {
            return new ParsedField
            {
                Name = _header,
                RawValue = null,
                DataType = XLDataType.Text,
                CellAddress = null,
                WorksheetName = worksheet.Name
            };
        }

        var cell = worksheet.Cell(rowNumber, _columnIndex.Value + 1);
        var rawValue = GetCellValue(cell);

        return new ParsedField
        {
            Name = _header,
            RawValue = rawValue,
            DataType = cell.DataType,
            CellAddress = cell.Address.ToString(),
            WorksheetName = worksheet.Name
        };
    }

    /// <summary>
    /// Получить значение ячейки как строку
    /// </summary>
    private string? GetCellValue(IXLCell cell)
    {
        if (cell.IsEmpty())
            return null;

        try
        {
            var value = cell.GetValue<string>();

            if (string.IsNullOrWhiteSpace(value))
                return null;

            return value;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Конвертировать значение поля в указанный тип
    /// </summary>
    public T? ConvertTo<T>(ParsedField field)
    {
        if (field.RawValue == null)
            return default;

        var targetType = typeof(T);
        var converted = _typeConverter.Convert(field.RawValue, field.DataType, targetType);

        if (converted == null)
            return default;

        return (T)converted;
    }

    /// <summary>
    /// Была ли найдена колонка
    /// </summary>
    public bool IsFound => _columnIndex.HasValue && _columnIndex.Value >= 0;
}
