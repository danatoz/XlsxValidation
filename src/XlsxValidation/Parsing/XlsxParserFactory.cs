using System.Collections.Concurrent;
using XlsxValidation.Configuration;

namespace XlsxValidation.Parsing;

/// <summary>
/// Фабрика для создания парсеров по имени профиля
/// </summary>
public class XlsxParserFactory
{
    private readonly ConcurrentDictionary<string, XlsxParser> _parsersCache = new();
    private readonly Dictionary<string, XlsxProfileConfig> _profiles;

    /// <summary>
    /// Создать фабрику парсеров
    /// </summary>
    public XlsxParserFactory(Dictionary<string, XlsxProfileConfig> profiles)
    {
        _profiles = profiles;
    }

    /// <summary>
    /// Создать парсер для профиля
    /// </summary>
    /// <param name="profileName">Имя профиля</param>
    /// <returns>Парсер для указанного профиля</returns>
    /// <exception cref="ProfileNotFoundException">Если профиль не найден</exception>
    public XlsxParser CreateForProfile(string profileName)
    {
        return _parsersCache.GetOrAdd(profileName, name =>
        {
            if (!_profiles.TryGetValue(name, out var config))
                throw new ProfileNotFoundException(name);

            var typeConverter = new TypeConverter(config.Parsing.Options);
            return XlsxParser.FromConfig(name, config, typeConverter);
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
    /// Очистить кэш парсеров
    /// </summary>
    public void ClearCache()
    {
        _parsersCache.Clear();
    }

    /// <summary>
    /// Очистить кэш для конкретного профиля
    /// </summary>
    public void ClearCacheForProfile(string profileName)
    {
        _parsersCache.TryRemove(profileName, out _);
    }
}

/// <summary>
/// Исключение: профиль не найден
/// </summary>
public class ProfileNotFoundException : Exception
{
    /// <summary>
    /// Имя профиля
    /// </summary>
    public string ProfileName { get; }

    public ProfileNotFoundException(string profileName)
        : base($"Профиль '{profileName}' не найден")
    {
        ProfileName = profileName;
    }

    public ProfileNotFoundException(string profileName, string message)
        : base(message)
    {
        ProfileName = profileName;
    }

    public ProfileNotFoundException(string profileName, string message, Exception innerException)
        : base(message, innerException)
    {
        ProfileName = profileName;
    }
}
