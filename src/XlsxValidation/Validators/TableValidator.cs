using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;
using XlsxValidation.Results;
using XlsxValidation.Rules;

namespace XlsxValidation.Validators;

/// <summary>
/// Конфигурация правила для колонки таблицы
/// </summary>
public record ColumnRuleSet
{
    public string Header { get; init; } = string.Empty;
    public string FieldName { get; init; } = string.Empty;
    public List<Func<IXLCell, int, ValidationResult>> Rules { get; init; } = new();
}

/// <summary>
/// Валидатор таблицы
/// </summary>
public class TableValidator
{
    private readonly string _tableName;
    private readonly ICellAnchor _headerAnchor;
    private readonly StopConditionConfig? _stopCondition;
    private readonly int? _maxRows;
    private readonly List<ColumnRuleSet> _columnRules;

    public TableValidator(
        string tableName,
        ICellAnchor headerAnchor,
        StopConditionConfig? stopCondition,
        int? maxRows,
        List<ColumnRuleSet> columnRules)
    {
        _tableName = tableName;
        _headerAnchor = headerAnchor;
        _stopCondition = stopCondition;
        _maxRows = maxRows;
        _columnRules = columnRules;
    }

    /// <summary>
    /// Валидировать таблицу на листе
    /// </summary>
    public IEnumerable<ValidationError> Validate(IXLWorksheet worksheet, string? worksheetName = null)
    {
        var errors = new List<ValidationError>();

        // Найти строку заголовков
        var headerResult = _headerAnchor.Resolve(worksheet);
        if (!headerResult.IsSuccess || headerResult.Cell == null)
        {
            errors.Add(new ValidationError
            {
                FieldName = _tableName,
                CellAddress = null,
                RuleId = "AnchorNotFound",
                Message = $"Заголовок таблицы не найден: {headerResult.ErrorMessage}",
                WorksheetName = worksheetName
            });
            return errors;
        }

        var headerCell = headerResult.Cell!;
        var headerRow = headerCell.Address.RowNumber;

        // Построить маппинг колонок: заголовок → номер колонки
        var columnMapping = BuildColumnMapping(worksheet, headerRow);

        // Валидировать каждую колонку
        foreach (var columnRuleSet in _columnRules)
        {
            if (!columnMapping.TryGetValue(columnRuleSet.Header, out var columnNumber))
            {
                errors.Add(new ValidationError
                {
                    FieldName = $"{_tableName}.{columnRuleSet.FieldName}",
                    CellAddress = null,
                    RuleId = "ColumnNotFound",
                    Message = $"Колонка '{columnRuleSet.Header}' не найдена в заголовках",
                    WorksheetName = worksheetName
                });
                continue;
            }

            // Валидировать ячейки колонки
            var columnErrors = ValidateColumn(worksheet, columnNumber, columnRuleSet, headerRow + 1, worksheetName);
            errors.AddRange(columnErrors);
        }

        return errors;
    }

    /// <summary>
    /// Построить маппинг заголовков к номерам колонок
    /// </summary>
    private Dictionary<string, int> BuildColumnMapping(IXLWorksheet worksheet, int headerRow)
    {
        var mapping = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var usedRange = worksheet.RangeUsed();
        
        if (usedRange == null || headerRow < usedRange.FirstRow().RowNumber())
            return mapping;

        var maxColumn = usedRange.LastColumn().ColumnNumber();
        
        for (int col = 1; col <= maxColumn; col++)
        {
            var cell = worksheet.Cell(headerRow, col);
            var value = cell.GetValue<string>()?.Trim();
            
            if (!string.IsNullOrEmpty(value) && !mapping.ContainsKey(value))
                mapping[value] = col;
        }

        return mapping;
    }

    /// <summary>
    /// Валидировать колонку таблицы
    /// </summary>
    private IEnumerable<ValidationError> ValidateColumn(
        IXLWorksheet worksheet,
        int columnNumber,
        ColumnRuleSet columnRuleSet,
        int startRow,
        string? worksheetName)
    {
        var errors = new List<ValidationError>();
        var maxRows = _maxRows ?? int.MaxValue;
        var rowsProcessed = 0;

        for (int row = startRow; row <= XLHelper.MaxRowNumber; row++)
        {
            if (rowsProcessed >= maxRows)
                break;

            var cell = worksheet.Cell(row, columnNumber);

            // Проверка условия остановки
            if (ShouldStop(cell, row, startRow))
                break;

            rowsProcessed++;

            // Применить правила
            foreach (var rule in columnRuleSet.Rules)
            {
                var ruleResult = rule(cell, row);
                if (!ruleResult.IsValid && ruleResult.ErrorMessage != null)
                {
                    var ruleId = ExtractRuleId(ruleResult.ErrorMessage);
                    
                    errors.Add(new ValidationError
                    {
                        FieldName = $"{_tableName}.{columnRuleSet.FieldName}",
                        CellAddress = cell.Address.ToString(),
                        RowNumber = row,
                        RuleId = ruleId,
                        Message = ruleResult.ErrorMessage,
                        WorksheetName = worksheetName
                    });
                }
            }
        }

        return errors;
    }

    /// <summary>
    /// Проверить условие остановки итерации
    /// </summary>
    private bool ShouldStop(IXLCell cell, int currentRow, int startRow)
    {
        if (_stopCondition == null)
            return false;

        return _stopCondition.Type switch
        {
            StopConditionType.EmptyRow => IsEmptyRow(cell, currentRow),
            StopConditionType.SentinelValue => IsSentinelValue(cell),
            StopConditionType.MaxRows => false, // Обрабатывается в цикле
            _ => false
        };
    }

    /// <summary>
    /// Проверить, является ли строка пустой
    /// </summary>
    private bool IsEmptyRow(IXLCell cell, int row)
    {
        // Проверяем всю строку начиная с первой колонки
        var worksheet = cell.Worksheet;
        var maxColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 1;
        
        for (int col = 1; col <= maxColumn; col++)
        {
            var c = worksheet.Cell(row, col);
            if (!c.IsEmpty())
                return false;
        }
        
        return true;
    }

    /// <summary>
    /// Проверить значение-маркер остановки
    /// </summary>
    private bool IsSentinelValue(IXLCell cell)
    {
        if (_stopCondition?.SentinelValue == null)
            return false;

        var value = cell.GetValue<string>()?.Trim();
        return value == _stopCondition.SentinelValue;
    }

    /// <summary>
    /// Извлечь идентификатор правила из сообщения
    /// </summary>
    private string ExtractRuleId(string errorMessage)
    {
        if (errorMessage.StartsWith("["))
        {
            var endBracket = errorMessage.IndexOf(']');
            if (endBracket > 0)
                return errorMessage[1..endBracket];
        }
        return "Unknown";
    }
}
