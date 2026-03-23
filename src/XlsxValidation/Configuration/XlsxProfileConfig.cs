namespace XlsxValidation.Configuration;

/// <summary>
/// Конфигурация валидации одиночной ячейки
/// </summary>
public record CellValidationConfig
{
    /// <summary>
    /// Логическое имя поля
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Якорь для поиска ячейки
    /// </summary>
    public AnchorConfig Anchor { get; init; } = new();

    /// <summary>
    /// Список правил валидации
    /// </summary>
    public List<RuleConfig> Rules { get; init; } = new();
}

/// <summary>
/// Конфигурация колонки таблицы
/// </summary>
public record ColumnConfig
{
    /// <summary>
    /// Заголовок колонки (для поиска в заголовках таблицы)
    /// </summary>
    public string Header { get; init; } = string.Empty;

    /// <summary>
    /// Логическое имя поля (опционально, если отличается от заголовка)
    /// </summary>
    public string? Name { get; init; }

    /// <summary>
    /// Список правил валидации для колонки
    /// </summary>
    public List<RuleConfig> Rules { get; init; } = new();
}

/// <summary>
/// Тип условия остановки итерации по таблице
/// </summary>
public enum StopConditionType
{
    EmptyRow,
    SentinelValue,
    MaxRows
}

/// <summary>
/// Конфигурация условия остановки итерации
/// </summary>
public record StopConditionConfig
{
    /// <summary>
    /// Тип условия остановки
    /// </summary>
    public StopConditionType Type { get; init; }

    /// <summary>
    /// Значение-маркер для SentinelValue
    /// </summary>
    public string? SentinelValue { get; init; }
}

/// <summary>
/// Конфигурация валидации таблицы
/// </summary>
public record TableValidationConfig
{
    /// <summary>
    /// Логическое имя таблицы
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Якорь для поиска строки заголовков
    /// </summary>
    public AnchorConfig HeaderAnchor { get; init; } = new();

    /// <summary>
    /// Условие остановки итерации
    /// </summary>
    public StopConditionConfig? StopCondition { get; init; }

    /// <summary>
    /// Максимальное количество строк для валидации
    /// </summary>
    public int? MaxRows { get; init; }

    /// <summary>
    /// Конфигурация колонок
    /// </summary>
    public List<ColumnConfig> Columns { get; init; } = new();
}

/// <summary>
/// Конфигурация валидации листа
/// </summary>
public record WorksheetValidationConfig
{
    /// <summary>
    /// Имя листа (или null для первого листа)
    /// </summary>
    public string? Name { get; init; }

    /// <summary>
    /// Валидация одиночных ячеек
    /// </summary>
    public List<CellValidationConfig> Cells { get; init; } = new();

    /// <summary>
    /// Валидация таблиц
    /// </summary>
    public List<TableValidationConfig> Tables { get; init; } = new();
}

/// <summary>
/// Корневая конфигурация профиля валидации
/// </summary>
public record XlsxProfileConfig
{
    /// <summary>
    /// Имя профиля (идентификатор)
    /// </summary>
    public string Profile { get; init; } = string.Empty;

    /// <summary>
    /// Описание профиля
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Версия профиля
    /// </summary>
    public string? Version { get; init; }

    /// <summary>
    /// Конфигурация валидации
    /// </summary>
    public ProfileValidationSection Validation { get; init; } = new();

    /// <summary>
    /// Служебный блок для YAML-якорей (игнорируется при десериализации)
    /// </summary>
    public Dictionary<string, object>? Rules { get; init; }
}

/// <summary>
/// Секция валидации в профиле
/// </summary>
public record ProfileValidationSection
{
    /// <summary>
    /// Конфигурация валидации по листам
    /// </summary>
    public List<WorksheetValidationConfig> Worksheets { get; init; } = new();
}
