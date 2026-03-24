using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;

namespace XlsxValidation.Parsing;

/// <summary>
/// Парсер XLSX файлов
/// </summary>
public class XlsxParser : IXlsxParser
{
    private readonly string _profileName;
    private readonly ParsingSection _parsingConfig;
    private readonly List<CellParser> _cellParsers;
    private readonly List<TableParser> _tableParsers;
    private readonly TypeConverter _typeConverter;

    /// <summary>
    /// Создать парсер
    /// </summary>
    public XlsxParser(
        string profileName,
        ParsingSection parsingConfig,
        List<CellParser> cellParsers,
        List<TableParser> tableParsers)
    {
        _profileName = profileName;
        _parsingConfig = parsingConfig;
        _cellParsers = cellParsers;
        _tableParsers = tableParsers;
        _typeConverter = new TypeConverter(parsingConfig.Options);
    }

    /// <summary>
    /// Распарсить XLSX файл из потока
    /// </summary>
    public XlsxParseResult Parse(Stream stream)
    {
        using var workbook = new XLWorkbook(stream);
        return Parse(workbook);
    }

    /// <summary>
    /// Распарсить XLSX файл из пути
    /// </summary>
    public XlsxParseResult Parse(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        return Parse(workbook);
    }

    /// <summary>
    /// Распарсить книгу ClosedXML
    /// </summary>
    public XlsxParseResult Parse(IXLWorkbook workbook)
    {
        var errors = new List<ParseError>();
        var fields = new List<ParsedField>();
        var tables = new List<ParsedTable>();

        // Парсинг одиночных ячеек (используем первый лист по умолчанию)
        var worksheet = workbook.Worksheets.First();
        var worksheetName = worksheet.Name;

        foreach (var cellParser in _cellParsers)
        {
            try
            {
                var field = cellParser.Parse(worksheet, worksheetName);
                fields.Add(field);
            }
            catch (Exception ex)
            {
                errors.Add(ParseError.Create(
                    cellParser.ToString() ?? "unknown",
                    $"Ошибка парсинга ячейки: {ex.Message}",
                    exception: ex,
                    worksheetName: worksheetName));
            }
        }

        // Парсинг таблиц
        foreach (var tableParser in _tableParsers)
        {
            try
            {
                var table = tableParser.Parse(worksheet, worksheetName);
                tables.Add(table);
            }
            catch (Exception ex)
            {
                errors.Add(ParseError.Create(
                    tableParser.ToString() ?? "unknown",
                    $"Ошибка парсинга таблицы: {ex.Message}",
                    exception: ex,
                    worksheetName: worksheetName));
            }
        }

        if (errors.Count > 0)
        {
            return XlsxParseResult.WithErrors(_profileName, errors)
            with
            {
                Fields = fields,
                Tables = tables
            };
        }

        return XlsxParseResult.Success(_profileName, fields, tables);
    }

    /// <summary>
    /// Создать парсер из конфигурации
    /// </summary>
    public static XlsxParser FromConfig(
        string profileName,
        XlsxProfileConfig config,
        TypeConverter typeConverter)
    {
        var cellParsers = new List<CellParser>();
        var tableParsers = new List<TableParser>();

        // Создать парсеры для каждого листа
        foreach (var worksheetConfig in config.Validation.Worksheets)
        {
            // Парсеры одиночных ячеек
            foreach (var cellConfig in worksheetConfig.Cells)
            {
                var fieldType = config.Parsing.FieldTypes.TryGetValue(cellConfig.Name, out var type)
                    ? type
                    : "string";

                var anchor = AnchorFactory.CreateAnchor(cellConfig.Anchor);
                var cellParser = new CellParser(
                    cellConfig.Name,
                    anchor,
                    fieldType,
                    typeConverter);

                cellParsers.Add(cellParser);
            }

            // Парсеры таблиц
            foreach (var tableConfig in worksheetConfig.Tables)
            {
                var headerAnchor = AnchorFactory.CreateAnchor(tableConfig.HeaderAnchor);

                var columnParsers = tableConfig.Columns.Select(col =>
                {
                    var fieldType = config.Parsing.FieldTypes.TryGetValue(
                        $"{tableConfig.Name}.{col.Header}", out var type)
                        ? type
                        : "string";

                    return new ColumnParser(col.Header, fieldType, typeConverter);
                }).ToList();

                var tableParser = new TableParser(
                    tableConfig.Name,
                    headerAnchor,
                    tableConfig.StopCondition,
                    tableConfig.MaxRows,
                    columnParsers,
                    typeConverter);

                tableParsers.Add(tableParser);
            }
        }

        return new XlsxParser(
            profileName,
            config.Parsing,
            cellParsers,
            tableParsers);
    }
}
