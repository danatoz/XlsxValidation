namespace XlsxValidation.Configuration;

/// <summary>
/// Типы якорей для адресации ячеек
/// </summary>
public enum AnchorType
{
    Content,
    Offset,
    NamedRange,
    Address
}

/// <summary>
/// Конфигурация якоря для поиска ячейки
/// </summary>
public record AnchorConfig
{
    /// <summary>
    /// Тип якоря
    /// </summary>
    public AnchorType Type { get; init; }

    /// <summary>
    /// Значение якоря (для content, named-range, address)
    /// </summary>
    public string? Value { get; init; }

    /// <summary>
    /// Базовый якорь для offset
    /// </summary>
    public AnchorConfig? Base { get; init; }

    /// <summary>
    /// Смещение по строкам (для offset)
    /// </summary>
    public int RowOffset { get; init; }

    /// <summary>
    /// Смещение по колонкам (для offset)
    /// </summary>
    public int ColOffset { get; init; }

    /// <summary>
    /// Опция для content якоря - точное или частичное совпадение
    /// </summary>
    public bool ExactMatch { get; init; } = false;
}
