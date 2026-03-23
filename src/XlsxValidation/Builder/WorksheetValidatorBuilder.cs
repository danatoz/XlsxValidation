using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Validators;

namespace XlsxValidation.Builder;

/// <summary>
/// Builder для создания WorksheetValidator
/// </summary>
public class WorksheetValidatorBuilder
{
    private string? _worksheetName;
    private readonly List<CellValidator> _cellValidators = new();
    private readonly List<TableValidator> _tableValidators = new();
    private readonly XlsxRuleRegistry _registry;

    public WorksheetValidatorBuilder(XlsxRuleRegistry registry)
    {
        _registry = registry;
    }

    public WorksheetValidatorBuilder WithWorksheetName(string? name)
    {
        _worksheetName = name;
        return this;
    }

    public WorksheetValidatorBuilder AddCellValidator(CellValidator validator)
    {
        _cellValidators.Add(validator);
        return this;
    }

    public WorksheetValidatorBuilder AddTableValidator(TableValidator validator)
    {
        _tableValidators.Add(validator);
        return this;
    }

    public WorksheetValidatorBuilder FromConfig(WorksheetValidationConfig config)
    {
        _worksheetName = config.Name;

        // Создать CellValidator из конфига
        foreach (var cellConfig in config.Cells)
        {
            var cellBuilder = new CellValidatorBuilder(_registry);
            var cellValidator = cellBuilder
                .WithFieldName(cellConfig.Name)
                .WithAnchorFromConfig(cellConfig.Anchor)
                .AddRulesFromConfig(cellConfig.Rules)
                .Build();
            
            _cellValidators.Add(cellValidator);
        }

        // Создать TableValidator из конфига
        foreach (var tableConfig in config.Tables)
        {
            var tableBuilder = new TableValidatorBuilder(_registry);
            
            var stopCondition = tableConfig.StopCondition ?? new StopConditionConfig
            {
                Type = StopConditionType.EmptyRow
            };

            var tableValidator = tableBuilder
                .WithTableName(tableConfig.Name)
                .WithHeaderAnchorFromConfig(tableConfig.HeaderAnchor)
                .WithStopCondition(stopCondition)
                .WithMaxRows(tableConfig.MaxRows)
                .AddColumnsFromConfig(tableConfig.Columns)
                .Build();
            
            _tableValidators.Add(tableValidator);
        }

        return this;
    }

    public WorksheetValidator Build()
    {
        return new WorksheetValidator(_worksheetName, _cellValidators, _tableValidators);
    }
}
