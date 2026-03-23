namespace XlsxValidation.Results;

/// <summary>
/// Результат выполнения правила валидации
/// </summary>
public record ValidationResult
{
    /// <summary>
    /// Была ли валидация успешной
    /// </summary>
    public bool IsValid { get; init; }

    /// <summary>
    /// Сообщение об ошибке (если валидация не пройдена)
    /// </summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Создать успешный результат
    /// </summary>
    public static ValidationResult Ok() => new() { IsValid = true };

    /// <summary>
    /// Создать результат с ошибкой
    /// </summary>
    public static ValidationResult Error(string message) => new() 
    { 
        IsValid = false, 
        ErrorMessage = message 
    };
}

/// <summary>
/// Ошибка валидации
/// </summary>
public record ValidationError
{
    /// <summary>
    /// Логическое имя поля/колонки
    /// </summary>
    public string FieldName { get; init; } = string.Empty;

    /// <summary>
    /// Физический адрес ячейки (для отладки, nullable)
    /// </summary>
    public string? CellAddress { get; init; }

    /// <summary>
    /// Номер строки (для табличных ошибок, nullable)
    /// </summary>
    public int? RowNumber { get; init; }

    /// <summary>
    /// Идентификатор сработавшего правила
    /// </summary>
    public string RuleId { get; init; } = string.Empty;

    /// <summary>
    /// Человекочитаемое сообщение об ошибке
    /// </summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>
    /// Имя листа (опционально)
    /// </summary>
    public string? WorksheetName { get; init; }

    public override string ToString()
    {
        var location = CellAddress ?? (RowNumber.HasValue ? $"строка {RowNumber}" : "unknown");
        if (WorksheetName != null)
            location = $"{WorksheetName}!{location}";
        
        return $"[{RuleId}] {FieldName} ({location}): {Message}";
    }
}

/// <summary>
/// Предупреждение валидации (некритичная ошибка)
/// </summary>
public record ValidationWarning
{
    /// <summary>
    /// Логическое имя поля/колонки
    /// </summary>
    public string FieldName { get; init; } = string.Empty;

    /// <summary>
    /// Физический адрес ячейки (для отладки)
    /// </summary>
    public string? CellAddress { get; init; }

    /// <summary>
    /// Номер строки (для табличных предупреждений)
    /// </summary>
    public int? RowNumber { get; init; }

    /// <summary>
    /// Идентификатор правила
    /// </summary>
    public string RuleId { get; init; } = string.Empty;

    /// <summary>
    /// Человекочитаемое сообщение
    /// </summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>
    /// Имя листа (опционально)
    /// </summary>
    public string? WorksheetName { get; init; }
}

/// <summary>
/// Отчёт о валидации
/// </summary>
public record ValidationReport
{
    /// <summary>
    /// Была ли валидация успешной
    /// </summary>
    public bool IsValid => Errors.Count == 0;

    /// <summary>
    /// Имя профиля валидации
    /// </summary>
    public string ProfileName { get; init; } = string.Empty;

    /// <summary>
    /// Список ошибок
    /// </summary>
    public IReadOnlyList<ValidationError> Errors { get; init; } = new List<ValidationError>();

    /// <summary>
    /// Список предупреждений
    /// </summary>
    public IReadOnlyList<ValidationWarning> Warnings { get; init; } = new List<ValidationWarning>();

    /// <summary>
    /// Создать пустой отчёт
    /// </summary>
    public static ValidationReport Empty(string profileName) => new() 
    { 
        ProfileName = profileName,
        Errors = new List<ValidationError>(),
        Warnings = new List<ValidationWarning>()
    };

    /// <summary>
    /// Создать отчёт с ошибками
    /// </summary>
    public static ValidationReport WithErrors(string profileName, IEnumerable<ValidationError> errors) => new()
    {
        ProfileName = profileName,
        Errors = errors.ToList().AsReadOnly(),
        Warnings = new List<ValidationWarning>()
    };

    /// <summary>
    /// Создать успешный отчёт
    /// </summary>
    public static ValidationReport Success(string profileName) => new()
    {
        ProfileName = profileName,
        Errors = new List<ValidationError>(),
        Warnings = new List<ValidationWarning>()
    };
}
