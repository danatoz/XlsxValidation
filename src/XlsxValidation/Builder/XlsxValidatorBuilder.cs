using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Validators;

namespace XlsxValidation.Builder;

/// <summary>
/// Builder для создания XlsxValidator
/// </summary>
public class XlsxValidatorBuilder
{
    private string _profileName = string.Empty;
    private readonly List<WorksheetValidator> _worksheetValidators = new();
    private readonly XlsxRuleRegistry _registry;

    public XlsxValidatorBuilder(XlsxRuleRegistry registry)
    {
        _registry = registry;
    }

    public XlsxValidatorBuilder WithProfileName(string profileName)
    {
        _profileName = profileName;
        return this;
    }

    public XlsxValidatorBuilder AddWorksheetValidator(WorksheetValidator validator)
    {
        _worksheetValidators.Add(validator);
        return this;
    }

    public XlsxValidatorBuilder FromConfig(XlsxProfileConfig config)
    {
        _profileName = config.Profile;

        foreach (var worksheetConfig in config.Validation.Worksheets)
        {
            var worksheetBuilder = new WorksheetValidatorBuilder(_registry);
            var worksheetValidator = worksheetBuilder
                .WithWorksheetName(worksheetConfig.Name)
                .FromConfig(worksheetConfig)
                .Build();
            
            _worksheetValidators.Add(worksheetValidator);
        }

        return this;
    }

    public XlsxValidator Build()
    {
        if (string.IsNullOrEmpty(_profileName))
            throw new InvalidOperationException("ProfileName не установлен");

        return new XlsxValidator(_profileName, _worksheetValidators);
    }
}
