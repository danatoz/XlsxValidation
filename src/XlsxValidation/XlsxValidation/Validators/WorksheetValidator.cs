using ClosedXML.Excel;
using XlsxValidation.Results;

namespace XlsxValidation.Validators;

/// <summary>
/// Валидатор листа XLSX
/// </summary>
public class WorksheetValidator
{
    private readonly string? _worksheetName;
    private readonly List<CellValidator> _cellValidators;
    private readonly List<TableValidator> _tableValidators;

    public WorksheetValidator(
        string? worksheetName,
        List<CellValidator> cellValidators,
        List<TableValidator> tableValidators)
    {
        _worksheetName = worksheetName;
        _cellValidators = cellValidators;
        _tableValidators = tableValidators;
    }

    /// <summary>
    /// Валидировать лист книги
    /// </summary>
    public IEnumerable<ValidationError> Validate(IXLWorkbook workbook)
    {
        var errors = new List<ValidationError>();

        IXLWorksheet? worksheet;
        
        if (string.IsNullOrEmpty(_worksheetName))
        {
            // Первый лист
            worksheet = workbook.Worksheets.First();
        }
        else
        {
            worksheet = workbook.Worksheets.FirstOrDefault(w => 
                w.Name.Equals(_worksheetName, StringComparison.OrdinalIgnoreCase));
            
            if (worksheet == null)
            {
                errors.Add(new ValidationError
                {
                    FieldName = _worksheetName ?? "unknown",
                    CellAddress = null,
                    RuleId = "WorksheetNotFound",
                    Message = $"Лист '{_worksheetName}' не найден"
                });
                return errors;
            }
        }

        var worksheetName = worksheet.Name;

        // Валидировать ячейки
        foreach (var cellValidator in _cellValidators)
        {
            var cellErrors = cellValidator.Validate(worksheet, worksheetName);
            errors.AddRange(cellErrors);
        }

        // Валидировать таблицы
        foreach (var tableValidator in _tableValidators)
        {
            var tableErrors = tableValidator.Validate(worksheet, worksheetName);
            errors.AddRange(tableErrors);
        }

        return errors;
    }
}
