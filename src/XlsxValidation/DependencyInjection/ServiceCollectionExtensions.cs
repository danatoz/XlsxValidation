using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using XlsxValidation.Configuration;
using XlsxValidation.Factory;
using XlsxValidation.Parsing;
using XlsxValidation.Rules;
using XlsxValidation.Results;

namespace XlsxValidation.DependencyInjection;

/// <summary>
/// Опции для настройки валидации XLSX
/// </summary>
public class XlsxValidationOptions
{
    /// <summary>
    /// Директория с YAML-профилями
    /// </summary>
    public string ProfilesDirectory { get; set; } = "xlsx-profiles";

    /// <summary>
    /// Загружать ли профили при старте
    /// </summary>
    public bool LoadProfilesAtStartup { get; set; } = true;

    /// <summary>
    /// Включить ли функциональность парсинга
    /// </summary>
    public bool EnableParsing { get; set; } = false;
}

/// <summary>
/// Extension методы для регистрации сервисов валидации XLSX
/// </summary>
public static class ServiceCollectionExtensions
{
    /// <summary>
    /// Добавить сервисы валидации XLSX
    /// </summary>
    public static IServiceCollection AddXlsxValidation(
        this IServiceCollection services,
        Action<XlsxValidationOptions>? configureOptions = null)
    {
        var options = new XlsxValidationOptions();
        configureOptions?.Invoke(options);

        // Зарегистрировать реестр правил
        var registry = new XlsxRuleRegistry();
        BuiltInRules.RegisterDefaults(registry);
        services.AddSingleton(registry);

        // Загрузить профили
        Dictionary<string, XlsxProfileConfig> profiles = new();
        
        if (options.LoadProfilesAtStartup)
        {
            if (!string.IsNullOrEmpty(options.ProfilesDirectory))
            {
                var loader = new YamlProfileLoader();
                profiles = loader.LoadDirectory(options.ProfilesDirectory);
            }
        }

        services.AddSingleton(profiles);

        // Зарегистрировать фабрику валидаторов
        services.AddSingleton<XlsxValidatorFactory>();

        // Зарегистрировать загрузчик профилей
        services.AddSingleton<YamlProfileLoader>();

        // Зарегистрировать фабрику парсеров (если включено)
        if (options.EnableParsing)
        {
            services.AddSingleton<XlsxParserFactory>();
        }

        return services;
    }

    /// <summary>
    /// Добавить кастомное правило валидации
    /// </summary>
    /// <param name="services">Коллекция сервисов</param>
    /// <param name="ruleId">Идентификатор правила</param>
    /// <param name="factory">Фабрика правила</param>
    public static IServiceCollection AddCustomRule(
        this IServiceCollection services,
        string ruleId,
        Func<RuleConfig, string, Func<IXLCell, ValidationResult>> factory)
    {
        // Создаём новый реестр и регистрируем правило
        // Примечание: этот метод должен вызываться до AddXlsxValidation
        // или реестр должен быть обновлён в AddXlsxValidation
        var registry = new XlsxRuleRegistry();
        registry.RegisterSharedRule(ruleId, factory);
        services.AddSingleton(registry);

        return services;
    }

    /// <summary>
    /// Добавить кастомное правило валидации (упрощённая версия)
    /// </summary>
    public static IServiceCollection AddCustomRule(
        this IServiceCollection services,
        string ruleId,
        Func<IXLCell, ValidationResult> rule)
    {
        return services.AddCustomRule(ruleId, (_, _) => cell => rule(cell));
    }
}
