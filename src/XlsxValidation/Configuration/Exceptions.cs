namespace XlsxValidation.Configuration;

/// <summary>
/// Исключение, возникающее при загрузке некорректного YAML-профиля
/// </summary>
public class ProfileLoadException : Exception
{
    public string FileName { get; }
    public int? LineNumber { get; }

    public ProfileLoadException(string fileName, string message)
        : base(message)
    {
        FileName = fileName;
    }

    public ProfileLoadException(string fileName, int lineNumber, string message)
        : base(message)
    {
        FileName = fileName;
        LineNumber = lineNumber;
    }

    public ProfileLoadException(string fileName, string message, Exception innerException)
        : base(message, innerException)
    {
        FileName = fileName;
    }
}

/// <summary>
/// Исключение, возникающее при запросе несуществующего профиля
/// </summary>
public class ProfileNotFoundException : Exception
{
    public string ProfileName { get; }

    public ProfileNotFoundException(string profileName)
        : base($"Профиль '{profileName}' не найден")
    {
        ProfileName = profileName;
    }
}
