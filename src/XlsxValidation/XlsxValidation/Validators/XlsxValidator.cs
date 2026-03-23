using ClosedXML.Excel;
using XlsxValidation.Results;

namespace XlsxValidation.Validators;

/// <summary>
/// Главный валидатор XLSX файлов
/// </summary>
public class XlsxValidator
{
    private readonly string _profileName;
    private readonly List<WorksheetValidator> _worksheetValidators;

    public XlsxValidator(
        string profileName,
        List<WorksheetValidator> worksheetValidators)
    {
        _profileName = profileName;
        _worksheetValidators = worksheetValidators;
    }

    /// <summary>
    /// Валидировать XLSX файл из потока
    /// </summary>
    public ValidationReport Validate(Stream stream)
    {
        using var workbook = new XLWorkbook(stream);
        return Validate(workbook);
    }

    /// <summary>
    /// Валидировать XLSX файл из пути
    /// </summary>
    public ValidationReport Validate(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        return Validate(workbook);
    }

    /// <summary>
    /// Валидировать книгу ClosedXML
    /// </summary>
    public ValidationReport Validate(IXLWorkbook workbook)
    {
        var allErrors = new List<ValidationError>();

        foreach (var worksheetValidator in _worksheetValidators)
        {
            var errors = worksheetValidator.Validate(workbook);
            allErrors.AddRange(errors);
        }

        return ValidationReport.WithErrors(_profileName, allErrors);
    }

    /// <summary>
    /// Валидировать и выбросить исключение при ошибках
    /// </summary>
    /// <exception cref="XlsxValidationException">Если валидация не пройдена</exception>
    public void ValidateAndThrow(Stream stream)
    {
        using var workbook = new XLWorkbook(stream);
        ValidateAndThrow(workbook);
    }

    /// <summary>
    /// Валидировать и выбросить исключение при ошибках
    /// </summary>
    /// <exception cref="XlsxValidationException">Если валидация не пройдена</exception>
    public void ValidateAndThrow(string filePath)
    {
        using var workbook = new XLWorkbook(filePath);
        ValidateAndThrow(workbook);
    }

    /// <summary>
    /// Валидировать и выбросить исключение при ошибках
    /// </summary>
    /// <exception cref="XlsxValidationException">Если валидация не пройдена</exception>
    public void ValidateAndThrow(IXLWorkbook workbook)
    {
        var report = Validate(workbook);
        
        if (!report.IsValid)
        {
            throw new XlsxValidationException(report);
        }
    }
}

/// <summary>
/// Исключение валидации XLSX
/// </summary>
public class XlsxValidationException : Exception
{
    public ValidationReport Report { get; }

    public XlsxValidationException(ValidationReport report)
        : base($"Валидация не пройдена: {report.Errors.Count} ошибок")
    {
        Report = report;
    }

    public XlsxValidationException(ValidationReport report, string message)
        : base(message)
    {
        Report = report;
    }

    public XlsxValidationException(ValidationReport report, string message, Exception innerException)
        : base(message, innerException)
    {
        Report = report;
    }
}
