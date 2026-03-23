using ClosedXML.Excel;
using XlsxValidation.Configuration;
using XlsxValidation.Results;

namespace XlsxValidation.Rules;

/// <summary>
/// Делегат фабрики правила для ячейки
/// </summary>
/// <param name="config">Конфигурация правила</param>
/// <returns>Функция валидации ячейки</returns>
public delegate Func<IXLCell, ValidationResult> CellRuleFactory(RuleConfig config);

/// <summary>
/// Делегат фабрики правила для колонки таблицы
/// </summary>
/// <param name="config">Конфигурация правила</param>
/// <returns>Функция валидации ячейки с номером строки</returns>
public delegate Func<IXLCell, int, ValidationResult> ColumnRuleFactory(RuleConfig config);

/// <summary>
/// Реестр правил валидации
/// </summary>
public class XlsxRuleRegistry
{
    private readonly Dictionary<string, CellRuleFactory> _cellRules = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, ColumnRuleFactory> _columnRules = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Зарегистрировать правило для ячейки
    /// </summary>
    public void RegisterCellRule(string ruleId, CellRuleFactory factory)
    {
        if (_cellRules.ContainsKey(ruleId))
            throw new InvalidOperationException($"Правило '{ruleId}' уже зарегистрировано");
        
        _cellRules[ruleId] = factory;
    }

    /// <summary>
    /// Зарегистрировать правило для колонки таблицы
    /// </summary>
    public void RegisterColumnRule(string ruleId, ColumnRuleFactory factory)
    {
        if (_columnRules.ContainsKey(ruleId))
            throw new InvalidOperationException($"Правило '{ruleId}' уже зарегистрировано");
        
        _columnRules[ruleId] = factory;
    }

    /// <summary>
    /// Зарегистрировать универсальное правило (и для ячейки, и для колонки)
    /// </summary>
    /// <param name="ruleId">Идентификатор правила</param>
    /// <param name="factory">Фабрика правила, принимающая конфигурацию и префикс контекста</param>
    public void RegisterSharedRule(string ruleId, Func<RuleConfig, string, Func<IXLCell, ValidationResult>> factory)
    {
        // Регистрируем как cell rule
        RegisterCellRule(ruleId, config => factory(config, string.Empty));
        
        // Регистрируем как column rule с обёрткой
        RegisterColumnRule(ruleId, config => 
        {
            var cellRule = factory(config, "[строка {row}] ");
            return (cell, row) => cellRule(cell);
        });
    }

    /// <summary>
    /// Получить фабрику правила для ячейки
    /// </summary>
    public CellRuleFactory? GetCellRule(string ruleId)
    {
        return _cellRules.TryGetValue(ruleId, out var factory) ? factory : null;
    }

    /// <summary>
    /// Получить фабрику правила для колонки
    /// </summary>
    public ColumnRuleFactory? GetColumnRule(string ruleId)
    {
        return _columnRules.TryGetValue(ruleId, out var factory) ? factory : null;
    }

    /// <summary>
    /// Проверить, зарегистрировано ли правило
    /// </summary>
    public bool IsRegistered(string ruleId)
    {
        return _cellRules.ContainsKey(ruleId) || _columnRules.ContainsKey(ruleId);
    }

    /// <summary>
    /// Получить все зарегистрированные идентификаторы правил
    /// </summary>
    public IEnumerable<string> GetRegisteredRuleIds()
    {
        return _cellRules.Keys.Concat(_columnRules.Keys).Distinct(StringComparer.OrdinalIgnoreCase);
    }
}
