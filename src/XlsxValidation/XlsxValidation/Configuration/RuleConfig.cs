namespace XlsxValidation.Configuration;

/// <summary>
/// Конфигурация правила валидации
/// </summary>
public record RuleConfig
{
    /// <summary>
    /// Идентификатор правила (например, "not-empty", "is-numeric")
    /// </summary>
    public string Rule { get; init; } = string.Empty;

    /// <summary>
    /// Параметры правила (например, { max: 100 } для max-length)
    /// </summary>
    public Dictionary<string, object> Params { get; init; } = new();

    /// <summary>
    /// Условие выполнения правила (опционально)
    /// </summary>
    public ConditionalConfig? When { get; init; }

    /// <summary>
    /// Кастомное сообщение об ошибке (опционально)
    /// </summary>
    public string? Message { get; init; }
}

/// <summary>
/// Конфигурация условного выполнения правила
/// </summary>
public record ConditionalConfig
{
    /// <summary>
    /// Якорь для получения значения условия
    /// </summary>
    public AnchorConfig? Anchor { get; init; }

    /// <summary>
    /// Тип условия (equals, not-equals, contains, etc.)
    /// </summary>
    public string Condition { get; init; } = string.Empty;

    /// <summary>
    /// Значение для сравнения
    /// </summary>
    public object? Value { get; init; }
}
