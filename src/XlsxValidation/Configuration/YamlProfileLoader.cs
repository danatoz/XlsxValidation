using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace XlsxValidation.Configuration;

/// <summary>
/// Загрузчик YAML-профилей валидации
/// </summary>
public class YamlProfileLoader
{
    private readonly IDeserializer _deserializer;

    public YamlProfileLoader()
    {
        _deserializer = new DeserializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .Build();
    }

    /// <summary>
    /// Загрузить все профили из директории
    /// </summary>
    /// <param name="directory">Директория с YAML-файлами</param>
    /// <returns>Словарь: имя профиля → конфигурация</returns>
    public Dictionary<string, XlsxProfileConfig> LoadDirectory(string directory)
    {
        if (!Directory.Exists(directory))
            throw new DirectoryNotFoundException($"Директория профилей не найдена: {directory}");

        var profiles = new Dictionary<string, XlsxProfileConfig>();
        var yamlFiles = Directory.GetFiles(directory, "*.yaml", SearchOption.TopDirectoryOnly)
            .Concat(Directory.GetFiles(directory, "*.yml", SearchOption.TopDirectoryOnly));

        foreach (var filePath in yamlFiles)
        {
            var fileName = Path.GetFileName(filePath);
            
            // Файлы с префиксом _ считаются служебными
            if (fileName.StartsWith("_"))
                continue;

            try
            {
                var yaml = File.ReadAllText(filePath);
                var config = _deserializer.Deserialize<XlsxProfileConfig>(yaml);
                
                if (!string.IsNullOrEmpty(config.Profile))
                {
                    profiles[config.Profile] = config;
                }
            }
            catch (Exception ex)
            {
                throw new ProfileLoadException(fileName, $"Ошибка загрузки профиля: {ex.Message}", ex);
            }
        }

        return profiles;
    }

    /// <summary>
    /// Загрузить отдельный YAML-файл
    /// </summary>
    public XlsxProfileConfig LoadFile(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Файл профиля не найден: {filePath}");

        try
        {
            var yaml = File.ReadAllText(filePath);
            return _deserializer.Deserialize<XlsxProfileConfig>(yaml);
        }
        catch (Exception ex)
        {
            var fileName = Path.GetFileName(filePath);
            throw new ProfileLoadException(fileName, $"Ошибка загрузки профиля: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Загрузить YAML из строки
    /// </summary>
    public XlsxProfileConfig LoadYaml(string yaml)
    {
        try
        {
            return _deserializer.Deserialize<XlsxProfileConfig>(yaml);
        }
        catch (Exception ex)
        {
            throw new ProfileLoadException("<inline>", $"Ошибка парсинга YAML: {ex.Message}", ex);
        }
    }
}
