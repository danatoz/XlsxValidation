using System.Collections.Concurrent;
using XlsxValidation.Builder;
using XlsxValidation.Configuration;
using XlsxValidation.Rules;
using XlsxValidation.Validators;

namespace XlsxValidation.Factory;

/// <summary>
/// Фабрика для создания валидаторов по имени профиля
/// </summary>
public class XlsxValidatorFactory
{
    private readonly XlsxRuleRegistry _registry;
    private readonly YamlProfileLoader _profileLoader;
    private readonly ConcurrentDictionary<string, XlsxValidator> _validatorsCache = new();
    private readonly Dictionary<string, XlsxProfileConfig> _profiles;

    public XlsxValidatorFactory(
        XlsxRuleRegistry registry,
        Dictionary<string, XlsxProfileConfig> profiles)
    {
        _registry = registry;
        _profileLoader = new YamlProfileLoader();
        _profiles = profiles;
    }

    /// <summary>
    /// Создать валидатор для профиля
    /// </summary>
    /// <param name="profileName">Имя профиля</param>
    /// <exception cref="ProfileNotFoundException">Если профиль не найден</exception>
    public XlsxValidator CreateForProfile(string profileName)
    {
        return _validatorsCache.GetOrAdd(profileName, name =>
        {
            if (!_profiles.TryGetValue(name, out var config))
                throw new ProfileNotFoundException(name);

            var builder = new XlsxValidatorBuilder(_registry);
            return builder
                .WithProfileName(name)
                .FromConfig(config)
                .Build();
        });
    }

    /// <summary>
    /// Получить все доступные имена профилей
    /// </summary>
    public IEnumerable<string> GetAvailableProfiles()
    {
        return _profiles.Keys;
    }

    /// <summary>
    /// Проверить, существует ли профиль
    /// </summary>
    public bool HasProfile(string profileName)
    {
        return _profiles.ContainsKey(profileName);
    }

    /// <summary>
    /// Очистить кэш валидаторов
    /// </summary>
    public void ClearCache()
    {
        _validatorsCache.Clear();
    }

    /// <summary>
    /// Очистить кэш для конкретного профиля
    /// </summary>
    public void ClearCacheForProfile(string profileName)
    {
        _validatorsCache.TryRemove(profileName, out _);
    }
}
