using ClosedXML.Excel;

namespace XlsxValidation.Parsing;

/// <summary>
/// Интерфейс парсера XLSX файлов
/// </summary>
public interface IXlsxParser
{
    /// <summary>
    /// Распарсить XLSX файл из потока
    /// </summary>
    XlsxParseResult Parse(Stream stream);

    /// <summary>
    /// Распарсить XLSX файл из пути
    /// </summary>
    XlsxParseResult Parse(string filePath);

    /// <summary>
    /// Распарсить книгу ClosedXML
    /// </summary>
    XlsxParseResult Parse(IXLWorkbook workbook);
}
