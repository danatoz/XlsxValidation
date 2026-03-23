using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Results;
using XlsxValidation.Validators;

namespace XlsxValidation.Builder;

/// <summary>
/// Builder для создания CellValidator
/// </summary>
public class CellValidatorBuilder
{
    private string _fieldName = string.Empty;
    private ICellAnchor? _anchor;
    private readonly List<Func<IXLCell, ValidationResult>> _rules = new();
    private readonly List<ConditionalConfig> _conditions = new();
    private readonly XlsxRuleRegistry _registry;
    private readonly AnchorFactory _anchorFactory;

    public CellValidatorBuilder(XlsxRuleRegistry registry)
    {
        _registry = registry;
        _anchorFactory = new AnchorFactory();
    }

    public CellValidatorBuilder WithFieldName(string fieldName)
    {
        _fieldName = fieldName;
        return this;
    }

    public CellValidatorBuilder WithAnchor(ICellAnchor anchor)
    {
        _anchor = anchor;
        return this;
    }

    public CellValidatorBuilder WithAnchorFromConfig(AnchorConfig config)
    {
        _anchor = _anchorFactory.Create(config);
        return this;
    }

    public CellValidatorBuilder AddRule(Func<IXLCell, ValidationResult> rule)
    {
        _rules.Add(rule);
        return this;
    }

    public CellValidatorBuilder AddRuleFromConfig(RuleConfig config)
    {
        var factory = _registry.GetCellRule(config.Rule);
        if (factory == null)
            throw new InvalidOperationException($"Правило '{config.Rule}' не зарегистрировано");

        var rule = factory(config);
        _rules.Add(rule);

        if (config.When != null)
            _conditions.Add(config.When);

        return this;
    }

    public CellValidatorBuilder AddRulesFromConfig(IEnumerable<RuleConfig> configs)
    {
        foreach (var config in configs)
        {
            AddRuleFromConfig(config);
        }
        return this;
    }

    public CellValidator Build()
    {
        if (string.IsNullOrEmpty(_fieldName))
            throw new InvalidOperationException("FieldName не установлен");

        if (_anchor == null)
            throw new InvalidOperationException("Якорь не установлен");

        return new CellValidator(_fieldName, _anchor, _rules, _conditions);
    }
}
