using XlsxValidation.Configuration;

namespace XlsxValidation.Anchors;

/// <summary>
/// Фабрика для создания якорей из конфигурации
/// </summary>
public class AnchorFactory
{
    /// <summary>
    /// Создать якорь из конфигурации
    /// </summary>
    public ICellAnchor Create(AnchorConfig config)
    {
        switch (config.Type)
        {
            case AnchorType.Content:
                if (string.IsNullOrEmpty(config.Value))
                    throw new ArgumentException("Value required for content anchor", nameof(config));
                return new ContentAnchor(config.Value, config.ExactMatch, null);

            case AnchorType.Offset:
                if (config.Base == null)
                    throw new ArgumentException("Base anchor required for offset anchor", nameof(config));
                var baseAnchor = Create(config.Base);
                return new OffsetAnchor(baseAnchor, config.RowOffset, config.ColOffset);

            case AnchorType.NamedRange:
                if (string.IsNullOrEmpty(config.Value))
                    throw new ArgumentException("Value required for named-range anchor", nameof(config));
                return new NamedRangeAnchor(config.Value);

            case AnchorType.Address:
                if (string.IsNullOrEmpty(config.Value))
                    throw new ArgumentException("Value required for address anchor", nameof(config));
                return new AddressAnchor(config.Value);

            default:
                throw new ArgumentException($"Неизвестный тип якоря: {config.Type}", nameof(config));
        }
    }

    /// <summary>
    /// Создать якорь из строкового представления типа
    /// </summary>
    public static AnchorType ParseAnchorType(string type)
    {
        switch (type.ToLowerInvariant())
        {
            case "content":
                return AnchorType.Content;
            case "offset":
                return AnchorType.Offset;
            case "named-range":
                return AnchorType.NamedRange;
            case "address":
                return AnchorType.Address;
            default:
                throw new ArgumentException($"Неизвестный тип якоря: {type}");
        }
    }
}
