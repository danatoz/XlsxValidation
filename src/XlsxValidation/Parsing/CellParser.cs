using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;

namespace XlsxValidation.Parsing;

/// <summary>
/// Парсер одиночной ячейки
/// </summary>
public class CellParser
{
    private readonly string _fieldName;
    private readonly ICellAnchor _anchor;
    private readonly string _fieldType;
    private readonly TypeConverter _typeConverter;

    /// <summary>
    /// Создать парсер ячейки
    /// </summary>
    public CellParser(
        string fieldName,
        ICellAnchor anchor,
        string fieldType,
        TypeConverter typeConverter)
    {
        _fieldName = fieldName;
        _anchor = anchor;
        _fieldType = fieldType;
        _typeConverter = typeConverter;
    }

    /// <summary>
    /// Распарсить ячейку из листа
    /// </summary>
    public ParsedField Parse(IXLWorksheet worksheet, string worksheetName)
    {
        var result = _anchor.Resolve(worksheet);

        if (!result.IsSuccess || result.Cell == null)
        {
            return new ParsedField
            {
                Name = _fieldName,
                RawValue = null,
                DataType = XLDataType.Text,
                CellAddress = null,
                WorksheetName = worksheetName
            };
        }

        var cell = result.Cell;
        var rawValue = GetCellValue(cell);

        return new ParsedField
        {
            Name = _fieldName,
            RawValue = rawValue,
            DataType = cell.DataType,
            CellAddress = cell.Address.ToString(),
            WorksheetName = worksheetName
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
            // Для дат и чисел используем Value, чтобы получить сырое значение
            if (cell.DataType == XLDataType.DateTime)
            {
                var dateValue = cell.GetValue<DateTime>();
                return dateValue.ToString("yyyy-MM-dd");
            }

            var value = cell.GetValue<string>();

            if (string.IsNullOrWhiteSpace(value))
                return null;

            return value;
        }
        catch
        {
            return cell.Address.ToString();
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
}
