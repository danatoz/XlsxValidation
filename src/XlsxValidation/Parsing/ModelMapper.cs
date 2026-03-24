using System.Reflection;

namespace XlsxValidation.Parsing;

/// <summary>
/// Атрибут для маппинга свойства на поле XLSX
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class XlsxFieldAttribute : Attribute
{
    /// <summary>
    /// Имя поля в XLSX (как указано в профиле)
    /// </summary>
    public string Name { get; init; } = string.Empty;

    /// <summary>
    /// Имя таблицы (для свойств, которые маппятся на таблицы)
    /// </summary>
    public string? Table { get; init; }
}

/// <summary>
/// Атрибут для маппинга свойства на колонку таблицы
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class XlsxColumnAttribute : Attribute
{
    /// <summary>
    /// Заголовок колонки в таблице
    /// </summary>
    public string Header { get; init; } = string.Empty;
}

/// <summary>
/// Маппер результатов парсинга на domain-модели
/// </summary>
public static class ModelMapper
{
    /// <summary>
    /// Распарсить результат в модель типа T
    /// </summary>
    public static T MapTo<T>(this XlsxParseResult result) where T : new()
    {
        var model = new T();
        MapTo(result, model, typeof(T));
        return model;
    }

    /// <summary>
    /// Распарсить результат в существующую модель
    /// </summary>
    public static void MapTo<T>(this XlsxParseResult result, T model)
    {
        MapTo(result, model, typeof(T));
    }

    /// <summary>
    /// Распарсить результат в модель указанного типа
    /// </summary>
    public static object MapTo(this XlsxParseResult result, Type modelType)
    {
        var model = Activator.CreateInstance(modelType)
            ?? throw new InvalidOperationException($"Не удалось создать экземпляр типа {modelType.Name}");

        MapTo(result, model, modelType);
        return model;
    }

    /// <summary>
    /// Распарсить результат в модель
    /// </summary>
    private static void MapTo<T>(XlsxParseResult result, T model, Type modelType)
    {
        var properties = modelType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

        foreach (var property in properties)
        {
            if (!property.CanWrite)
                continue;

            // Проверить атрибут XlsxField
            var fieldAttr = property.GetCustomAttribute<XlsxFieldAttribute>();
            if (fieldAttr != null)
            {
                if (!string.IsNullOrEmpty(fieldAttr.Table))
                {
                    // Маппинг таблицы
                    MapTableProperty(result, model, property, fieldAttr);
                }
                else
                {
                    // Маппинг одиночного поля
                    MapFieldProperty(result, model, property, fieldAttr);
                }
            }
            else
            {
                // Попытка авто-маппинга по имени свойства
                AutoMapProperty(result, model, property);
            }
        }
    }

    /// <summary>
    /// Маппинг свойства поля
    /// </summary>
    private static void MapFieldProperty<T>(XlsxParseResult result, T model, PropertyInfo property, XlsxFieldAttribute attr)
    {
        var field = result.GetField(attr.Name);
        if (field == null)
            return;

        var value = ConvertFieldToType(field, property.PropertyType);
        if (value != null)
        {
            property.SetValue(model, value);
        }
    }

    /// <summary>
    /// Маппинг свойства таблицы
    /// </summary>
    private static void MapTableProperty<T>(XlsxParseResult result, T model, PropertyInfo property, XlsxFieldAttribute attr)
    {
        var table = result.GetTable(attr.Table!);
        if (table == null)
            return;

        var propertyType = property.PropertyType;
        var elementType = GetCollectionElementType(propertyType);

        if (elementType == null)
            return;

        var listType = typeof(List<>).MakeGenericType(elementType);
        var list = Activator.CreateInstance(listType);

        var addMethod = listType.GetMethod("Add");

        foreach (var row in table.Rows)
        {
            var item = Activator.CreateInstance(elementType);
            MapTableRowToItem(row, item, elementType);
            addMethod?.Invoke(list, new[] { item });
        }

        property.SetValue(model, list);
    }

    /// <summary>
    /// Авто-маппинг свойства по имени
    /// </summary>
    private static void AutoMapProperty<T>(XlsxParseResult result, T model, PropertyInfo property)
    {
        // Поиск по имени свойства
        var field = result.Fields.FirstOrDefault(f =>
            f.Name.Equals(property.Name, StringComparison.OrdinalIgnoreCase));

        if (field != null)
        {
            var value = ConvertFieldToType(field, property.PropertyType);
            if (value != null)
            {
                property.SetValue(model, value);
            }
        }
    }

    /// <summary>
    /// Маппинг строки таблицы на элемент коллекции
    /// </summary>
    private static void MapTableRowToItem(ParsedTableRow row, object? item, Type itemType)
    {
        if (item == null)
            return;

        var properties = itemType.GetProperties(BindingFlags.Public | BindingFlags.Instance);

        foreach (var property in properties)
        {
            if (!property.CanWrite)
                continue;

            // Проверить атрибут XlsxColumn
            var columnAttr = property.GetCustomAttribute<XlsxColumnAttribute>();
            string headerName = columnAttr?.Header ?? property.Name;

            if (row.Fields.TryGetValue(headerName, out var field))
            {
                var value = ConvertFieldToType(field, property.PropertyType);
                if (value != null)
                {
                    property.SetValue(item, value);
                }
            }
        }
    }

    /// <summary>
    /// Конвертировать поле в указанный тип
    /// </summary>
    private static object? ConvertFieldToType(ParsedField field, Type targetType)
    {
        if (field.RawValue == null)
            return null;

        // Обработка nullable типов
        var underlyingType = Nullable.GetUnderlyingType(targetType);
        if (underlyingType != null)
            targetType = underlyingType;

        try
        {
            return field.AsType(targetType);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Получить тип элемента коллекции
    /// </summary>
    private static Type? GetCollectionElementType(Type collectionType)
    {
        if (collectionType.IsGenericType)
        {
            var genericType = collectionType.GetGenericTypeDefinition();

            if (genericType == typeof(List<>) ||
                genericType == typeof(IList<>) ||
                genericType == typeof(IEnumerable<>) ||
                genericType == typeof(ICollection<>))
            {
                return collectionType.GetGenericArguments()[0];
            }
        }

        // Поиск интерфейса IEnumerable<T>
        foreach (var iface in collectionType.GetInterfaces())
        {
            if (iface.IsGenericType && iface.GetGenericTypeDefinition() == typeof(IEnumerable<>))
            {
                return iface.GetGenericArguments()[0];
            }
        }

        return null;
    }
}
