# План реализации парсинга XLSX

## Обзор
Этот документ описывает пошаговый план реализации функциональности парсинга XLSX файлов.

## Фаза 1: Базовая инфраструктура (Week 1)

### Задача 1.1: Создать модели результатов парсинга
**Файл:** `src/XlsxValidation/Parsing/ParseResult.cs`

```csharp
public record ParsedField
{
    public string Name { get; init; }
    public string? RawValue { get; init; }
    public XLDataType DataType { get; init; }
    public string? CellAddress { get; init; }
    public string? WorksheetName { get; init; }
}

public record ParsedTableRow
{
    public int RowNumber { get; init; }
    public Dictionary<string, ParsedField> Fields { get; init; }
}

public record ParsedTable
{
    public string Name { get; init; }
    public IReadOnlyList<string> Headers { get; init; }
    public IReadOnlyList<ParsedTableRow> Rows { get; init; }
}

public record ParseError
{
    public string FieldName { get; init; }
    public string? CellAddress { get; init; }
    public int? RowNumber { get; init; }
    public string Message { get; init; }
    public Exception? Exception { get; init; }
}

public record XlsxParseResult
{
    public string ProfileName { get; init; }
    public IReadOnlyList<ParsedField> Fields { get; init; }
    public IReadOnlyList<ParsedTable> Tables { get; init; }
    public IReadOnlyList<ParseError> Errors { get; init; }
    public bool IsSuccess => Errors.Count == 0;
}
```

**Критерии приемки:**
- [ ] Все модели имеют init-only свойства
- [ ] Реализованы методы равенства (record)
- [ ] Добавлены XML-документы

---

### Задача 1.2: Создать конвертер типов
**Файл:** `src/XlsxValidation/Parsing/TypeConverter.cs`

```csharp
public class TypeConverter
{
    private readonly CultureInfo _culture;
    private readonly string[] _dateFormats;
    private readonly NumberStyles _numberStyles;

    public TypeConverter(ParseOptions options);

    public T? Convert<T>(string? value, XLDataType dataType);
    public string? ConvertToString(object? value);
    public decimal? ConvertToDecimal(string? value, XLDataType dataType);
    public int? ConvertToInt(string? value, XLDataType dataType);
    public DateTime? ConvertToDateTime(string? value, XLDataType dataType);
    public bool? ConvertToBoolean(string? value, XLDataType dataType);
}
```

**Критерии приемки:**
- [ ] Поддержка типов: string, int, long, decimal, double, DateTime, bool
- [ ] Обработка null и пустых значений
- [ ] Использование культуры из опций
- [ ] Unit-тесты для всех типов

---

### Задача 1.3: Создать конфигурацию парсинга
**Файл:** `src/XlsxValidation/Configuration/ParseConfig.cs`

```csharp
public record ParseOptions
{
    public bool SkipEmptyCells { get; init; } = true;
    public bool TrimStrings { get; init; } = true;
    public string Culture { get; init; } = "ru-RU";
    public string[] DateFormats { get; init; } = new[] { "dd.MM.yyyy" };
    public NumberStyles NumberStyles { get; init; } = NumberStyles.Number;
    public Dictionary<string, object> Defaults { get; init; } = new();
}

public record FieldMappingConfig
{
    public string Name { get; init; }
    public string Type { get; init; } = "string";
    public string? Column { get; init; }  // Для табличных полей
}

public record TableMappingConfig
{
    public string Name { get; init; }
    public string RowType { get; init; }
    public Dictionary<string, ColumnMappingConfig> Columns { get; init; }
}

public record ColumnMappingConfig
{
    public string Name { get; init; }
    public string Type { get; init; }
}

public record ParsingSection
{
    public Dictionary<string, string> FieldTypes { get; init; } = new();
    public ParseOptions Options { get; init; } = new();
    public Dictionary<string, TableMappingConfig> TableMapping { get; init; } = new();
}
```

**Критерии приемки:**
- [ ] Интеграция с YamlDotNet
- [ ] Обновление `XlsxProfileConfig` с секцией `Parsing`
- [ ] Unit-тесты на десериализацию YAML

---

## Фаза 2: Реализация парсера (Week 2)

### Задача 2.1: Создать интерфейс и базовый класс
**Файл:** `src/XlsxValidation/Parsing/IXlsxParser.cs`

```csharp
public interface IXlsxParser
{
    XlsxParseResult Parse(Stream stream);
    XlsxParseResult Parse(string filePath);
    XlsxParseResult Parse(IXLWorkbook workbook);
}
```

**Файл:** `src/XlsxValidation/Parsing/XlsxParser.cs`

```csharp
public class XlsxParser : IXlsxParser
{
    private readonly string _profileName;
    private readonly ParsingSection _parsingConfig;
    private readonly List<CellParser> _cellParsers;
    private readonly List<TableParser> _tableParsers;
    private readonly TypeConverter _typeConverter;

    // Реализация методов Parse
}
```

**Критерии приемки:**
- [ ] Интерфейс с тремя перегрузками Parse
- [ ] Конструктор принимает конфигурацию
- [ ] Зависимость на ClosedXML

---

### Задача 2.2: Создать парсер ячеек
**Файл:** `src/XlsxValidation/Parsing/CellParser.cs`

```csharp
public class CellParser
{
    private readonly string _fieldName;
    private readonly ICellAnchor _anchor;
    private readonly string _fieldType;
    private readonly TypeConverter _typeConverter;

    public ParsedField Parse(IXLWorksheet worksheet, string worksheetName);
}
```

**Критерии приемки:**
- [ ] Использование существующей системы якорей
- [ ] Конвертация значения в указанный тип
- [ ] Обработка ошибок (ячейка не найдена)

---

### Задача 2.3: Создать парсер таблиц
**Файл:** `src/XlsxValidation/Parsing/TableParser.cs`

```csharp
public class TableParser
{
    private readonly string _tableName;
    private readonly ICellAnchor _headerAnchor;
    private readonly StopConditionConfig _stopCondition;
    private readonly List<ColumnParser> _columnParsers;
    private readonly TypeConverter _typeConverter;

    public ParsedTable Parse(IXLWorksheet worksheet, string worksheetName);
}
```

**Критерии приемки:**
- [ ] Поиск заголовка через якорь
- [ ] Итерация по строкам до условия остановки
- [ ] Парсинг каждой колонки в строке

---

### Задача 2.4: Extension-методы для ParsedField
**Файл:** `src/XlsxValidation/Parsing/ParsedFieldExtensions.cs`

```csharp
public static class ParsedFieldExtensions
{
    public static string? AsString(this ParsedField field);
    public static int? AsInteger(this ParsedField field);
    public static long? AsLong(this ParsedField field);
    public static decimal? AsDecimal(this ParsedField field);
    public static double? AsDouble(this ParsedField field);
    public static DateTime? AsDateTime(this ParsedField field);
    public static bool? AsBoolean(this ParsedField field);
    public static T? AsType<T>(this ParsedField field);
}
```

**Критерии приемки:**
- [ ] Все методы возвращают nullable тип
- [ ] Обработка null значений
- [ ] Использование TypeConverter внутри

---

## Фаза 3: Интеграция (Week 3)

### Задача 3.1: Обновить XlsxProfileConfig
**Файл:** `src/XlsxValidation/Configuration/XlsxProfileConfig.cs`

Добавить свойство:
```csharp
public record XlsxProfileConfig
{
    // ... существующие свойства ...

    /// <summary>
    /// Конфигурация парсинга
    /// </summary>
    public ParsingSection Parsing { get; init; } = new();
}
```

**Критерии приемки:**
- [ ] YAML десериализация работает с новой секцией
- [ ] Обратная совместимость (профили без parsing работают)

---

### Задача 3.2: Создать фабрику парсеров
**Файл:** `src/XlsxValidation/Factory/XlsxParserFactory.cs`

```csharp
public class XlsxParserFactory
{
    private readonly ConcurrentDictionary<string, XlsxParser> _parsersCache = new();
    private readonly Dictionary<string, XlsxProfileConfig> _profiles;

    public XlsxParserFactory(Dictionary<string, XlsxProfileConfig> profiles);

    public XlsxParser CreateForProfile(string profileName);
    public IEnumerable<string> GetAvailableProfiles();
    public bool HasProfile(string profileName);
}
```

**Критерии приемки:**
- [ ] Кэширование парсеров
- [ ] Обработка отсутствующих профилей

---

### Задача 3.3: Обновить DI контейнер
**Файл:** `src/XlsxValidation/DependencyInjection/ServiceCollectionExtensions.cs`

Добавить опцию:
```csharp
public class XlsxValidationOptions
{
    // ...
    public bool EnableParsing { get; set; } = false;
}
```

Добавить регистрацию:
```csharp
public static IServiceCollection AddXlsxValidation(...)
{
    // ...
    if (options.EnableParsing)
    {
        services.AddSingleton<XlsxParserFactory>();
    }
}
```

**Критерии приемки:**
- [ ] Опциональная регистрация парсера
- [ ] Обратная совместимость

---

### Задача 3.4: Mapper на domain-модели
**Файл:** `src/XlsxValidation/Parsing/ModelMapper.cs`

```csharp
[AttributeUsage(AttributeTargets.Property)]
public class XlsxFieldAttribute : Attribute
{
    public string Name { get; init; }
    public string Table { get; init; }  // Для табличных полей
}

public static class ModelMapper
{
    public static T MapTo<T>(this XlsxParseResult result) where T : new();
    public static object MapTo(this XlsxParseResult result, Type modelType);
}
```

**Критерии приемки:**
- [ ] Поддержка атрибутов [XlsxField]
- [ ] Маппинг простых типов
- [ ] Маппинг коллекций (таблицы)

---

## Фаза 4: Тестирование и документация (Week 4)

### Задача 4.1: Unit-тесты
**Директория:** `tests/XlsxValidation.Tests/Parsing/`

```
Parsing/
├── TypeConverterTests.cs
├── CellParserTests.cs
├── TableParserTests.cs
├── ParsedFieldExtensionsTests.cs
├── ModelMapperTests.cs
└── XlsxParserIntegrationTests.cs
```

**Критерии приемки:**
- [ ] Покрытие > 80%
- [ ] Тесты на граничные случаи
- [ ] Тесты на ошибки

---

### Задача 4.2: Обновить README
**Файл:** `README.md` и `README.ru.md`

Добавить секции:
- Quick Start для парсинга
- Примеры использования
- Таблица методов конвертации

**Критерии приемки:**
- [ ] Пример кода в README
- [ ] Документация API
- [ ] Пример YAML профиля

---

### Задача 4.3: Примеры использования
**Директория:** `examples/`

```
examples/
├── InvoiceParsingExample.cs
├── SalaryReportParsingExample.cs
└── CustomMappingExample.cs
```

**Критерии приемки:**
- [ ] Рабочие примеры
- [ ] Комментарии в коде

---

## Итоговая структура проекта

```
src/XlsxValidation/
├── Parsing/
│   ├── IXlsxParser.cs
│   ├── XlsxParser.cs
│   ├── CellParser.cs
│   ├── TableParser.cs
│   ├── ColumnParser.cs
│   ├── ParseResult.cs              # ParsedField, ParsedTable, etc.
│   ├── TypeConverter.cs
│   ├── ModelMapper.cs              # XlsxFieldAttribute, MapTo<T>
│   └── ParsedFieldExtensions.cs    # AsString, AsDecimal, etc.
├── Configuration/
│   ├── XlsxProfileConfig.cs        # + ParsingSection
│   ├── ParseConfig.cs              # ParseOptions, FieldMappingConfig
│   ├── AnchorConfig.cs
│   ├── RuleConfig.cs
│   └── YamlProfileLoader.cs
├── Factory/
│   ├── XlsxValidatorFactory.cs
│   └── XlsxParserFactory.cs        # новый
├── DependencyInjection/
│   └── ServiceCollectionExtensions.cs  # + EnableParsing option
└── ... (существующие файлы)
```

## Метрики успеха

1. **Функциональность:**
   - [ ] Парсинг одиночных ячеек работает
   - [ ] Парсинг таблиц работает
   - [ ] Конвертация типов работает
   - [ ] MapTo<T>() работает

2. **Качество кода:**
   - [ ] Покрытие тестами > 80%
   - [ ] Нет breaking changes в существующем API
   - [ ] XML-документация на всех публичных API

3. **Документация:**
   - [ ] README обновлён
   - [ ] Примеры использования работают
   - [ ] ADR задокументирован

## Риски и зависимости

| Риск | Вероятность | Влияние | Митигация |
|------|-------------|---------|-----------|
| Сложность парсинга дат | Средняя | Средняя | Использовать стандартные форматы, документировать ограничения |
| Проблемы с кодировками | Низкая | Низкая | Использовать UTF-8, тестировать на разных файлах |
| Breaking changes в API | Низкая | Высокая | Тщательное ревью, семантическое версионирование |
| Производительность на больших файлах | Средняя | Средняя | Кэширование, ленивая загрузка, документировать лимиты |

## Timeline

- **Week 1:** Базовая инфраструктура (модели, конвертер, конфигурация)
- **Week 2:** Реализация парсера (CellParser, TableParser, XlsxParser)
- **Week 3:** Интеграция (DI, Factory, Mapper)
- **Week 4:** Тестирование и документация

**Итого:** 4 недели (20 рабочих дней)
