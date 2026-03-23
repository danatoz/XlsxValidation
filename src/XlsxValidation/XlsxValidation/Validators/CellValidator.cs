using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;
using XlsxValidation.Results;
using XlsxValidation.Rules;

namespace XlsxValidation.Validators;

/// <summary>
/// Валидатор одиночной ячейки
/// </summary>
public class CellValidator
{
    private readonly string _fieldName;
    private readonly ICellAnchor _anchor;
    private readonly List<Func<IXLCell, ValidationResult>> _rules;
    private readonly List<ConditionalConfig> _conditions;

    public CellValidator(
        string fieldName,
        ICellAnchor anchor,
        List<Func<IXLCell, ValidationResult>> rules,
        List<ConditionalConfig> conditions)
    {
        _fieldName = fieldName;
        _anchor = anchor;
        _rules = rules;
        _conditions = conditions;
    }

    /// <summary>
    /// Валидировать ячейку на листе
    /// </summary>
    public IEnumerable<ValidationError> Validate(IXLWorksheet worksheet, string? worksheetName = null)
    {
        var errors = new List<ValidationError>();

        // Разрешить якорь
        var result = _anchor.Resolve(worksheet);
        
        if (!result.IsSuccess)
        {
            errors.Add(new ValidationError
            {
                FieldName = _fieldName,
                CellAddress = null,
                RuleId = "AnchorNotFound",
                Message = $"Якорь не найден: {result.ErrorMessage}",
                WorksheetName = worksheetName
            });
            return errors;
        }

        var cell = result.Cell!;
        var cellAddress = cell.Address.ToString();

        // Проверка условий
        foreach (var condition in _conditions)
        {
            if (!CheckCondition(condition, worksheet, worksheetName))
                return errors; // Условие не выполнено - пропускаем валидацию
        }

        // Применить правила
        foreach (var rule in _rules)
        {
            var ruleResult = rule(cell);
            if (!ruleResult.IsValid && ruleResult.ErrorMessage != null)
            {
                // Извлекаем RuleId из сообщения (формат: "[префикс] сообщение")
                var ruleId = ExtractRuleId(ruleResult.ErrorMessage);
                
                errors.Add(new ValidationError
                {
                    FieldName = _fieldName,
                    CellAddress = cellAddress,
                    RuleId = ruleId,
                    Message = ruleResult.ErrorMessage,
                    WorksheetName = worksheetName
                });
            }
        }

        return errors;
    }

    /// <summary>
    /// Проверить условие выполнения правила
    /// </summary>
    private bool CheckCondition(ConditionalConfig condition, IXLWorksheet worksheet, string? worksheetName)
    {
        if (condition.Anchor == null)
            return true; // Нет якоря - условие считается выполненным

        var anchorFactory = new AnchorFactory();
        var anchor = anchorFactory.Create(condition.Anchor);
        var result = anchor.Resolve(worksheet);

        if (!result.IsSuccess || result.Cell == null)
            return false;

        var cellValue = result.Cell.GetValue<string>()?.Trim();
        var conditionValue = condition.Value?.ToString()?.Trim();

        return condition.Condition.ToLowerInvariant() switch
        {
            "equals" => cellValue == conditionValue,
            "not-equals" => cellValue != conditionValue,
            "contains" => cellValue?.Contains(conditionValue ?? "", StringComparison.OrdinalIgnoreCase) == true,
            "not-contains" => !(cellValue?.Contains(conditionValue ?? "", StringComparison.OrdinalIgnoreCase) == true),
            "not-empty" => !string.IsNullOrEmpty(cellValue),
            "empty" => string.IsNullOrEmpty(cellValue),
            _ => true
        };
    }

    /// <summary>
    /// Извлечь идентификатор правила из сообщения об ошибке
    /// </summary>
    private string ExtractRuleId(string errorMessage)
    {
        // Формат: "[префикс] сообщение" или просто "сообщение"
        if (errorMessage.StartsWith("["))
        {
            var endBracket = errorMessage.IndexOf(']');
            if (endBracket > 0)
                return errorMessage[1..endBracket];
        }
        return "Unknown";
    }
}
