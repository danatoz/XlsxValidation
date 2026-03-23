using ClosedXML.Excel;
using XlsxValidation.Anchors;
using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Results;
using XlsxValidation.Validators;

namespace XlsxValidation.Builder;

/// <summary>
/// Builder для создания TableValidator
/// </summary>
public class TableValidatorBuilder
{
    private string _tableName = string.Empty;
    private ICellAnchor? _headerAnchor;
    private StopConditionConfig? _stopCondition;
    private int? _maxRows;
    private readonly List<ColumnRuleSet> _columnRules = new();
    private readonly XlsxRuleRegistry _registry;
    private readonly AnchorFactory _anchorFactory;

    public TableValidatorBuilder(XlsxRuleRegistry registry)
    {
        _registry = registry;
        _anchorFactory = new AnchorFactory();
    }

    public TableValidatorBuilder WithTableName(string tableName)
    {
        _tableName = tableName;
        return this;
    }

    public TableValidatorBuilder WithHeaderAnchor(ICellAnchor anchor)
    {
        _headerAnchor = anchor;
        return this;
    }

    public TableValidatorBuilder WithHeaderAnchorFromConfig(AnchorConfig config)
    {
        _headerAnchor = _anchorFactory.Create(config);
        return this;
    }

    public TableValidatorBuilder WithStopCondition(StopConditionConfig? stopCondition)
    {
        _stopCondition = stopCondition;
        return this;
    }

    public TableValidatorBuilder WithMaxRows(int? maxRows)
    {
        _maxRows = maxRows;
        return this;
    }

    public TableValidatorBuilder AddColumn(ColumnRuleSet columnRuleSet)
    {
        _columnRules.Add(columnRuleSet);
        return this;
    }

    public TableValidatorBuilder AddColumnFromConfig(ColumnConfig config)
    {
        var fieldName = config.Name ?? config.Header;
        var rules = new List<Func<IXLCell, int, ValidationResult>>();

        foreach (var ruleConfig in config.Rules)
        {
            var factory = _registry.GetColumnRule(ruleConfig.Rule);
            if (factory == null)
                throw new InvalidOperationException($"Правило '{ruleConfig.Rule}' не зарегистрировано");

            var rule = factory(ruleConfig);
            rules.Add(rule);
        }

        _columnRules.Add(new ColumnRuleSet
        {
            Header = config.Header,
            FieldName = fieldName,
            Rules = rules
        });

        return this;
    }

    public TableValidatorBuilder AddColumnsFromConfig(IEnumerable<ColumnConfig> configs)
    {
        foreach (var config in configs)
        {
            AddColumnFromConfig(config);
        }
        return this;
    }

    public TableValidator Build()
    {
        if (string.IsNullOrEmpty(_tableName))
            throw new InvalidOperationException("TableName не установлен");

        if (_headerAnchor == null)
            throw new InvalidOperationException("Якорь заголовка не установлен");

        return new TableValidator(_tableName, _headerAnchor, _stopCondition, _maxRows, _columnRules);
    }
}
