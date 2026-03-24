# ADR-001: Добавление парсинга XLSX файлов с YAML-конфигурацией

## Статус
Предложен

## Контекст
Текущий проект **XlsxValidation** представляет собой библиотеку для **валидации** XLSX файлов с конфигурацией через YAML. Основная функциональность:
- Валидация ячеек и таблиц по правилам
- Система якорей для поиска ячеек (content, offset, named-range, address)
- 13 встроенных правил валидации
- Поддержка кастомных правил
- DI-интеграция

**Проблема**: Проект только валидирует данные, но не предоставляет возможности **извлечения** (парсинга) данных из XLSX файлов в структурированную модель. Пользователям необходимо вручную писать код для чтения тех же файлов, дублируя логику поиска ячеек через якоря.

**Цель**: Расширить библиотеку функциональностью парсинга XLSX файлов в strongly-typed модели с конфигурацией через YAML profile, сохраняя существующую архитектуру и систему якорей.

## Движущие силы
1. **Повторное использование системы якорей** — логика поиска ячеек уже реализована, её можно использовать для парсинга
2. **Единая конфигурация** — пользователи хотят описывать и валидацию, и парсинг в одном YAML-файле
3. **Type-safe извлечение данных** — необходимость получать типизированные модели вместо сырых значений
4. **Минимизация breaking changes** — новая функциональность не должна ломать существующий API валидации

## Предлагаемое решение

### 1. Модель данных для парсинга

#### Вариант A: Динамическая модель (Dictionary-based)
```csharp
public record ParsedWorksheet
{
    public string Name { get; init; } = string.Empty;
    public Dictionary<string, CellValue> Cells { get; init; } = new();
    public Dictionary<string, List<Dictionary<string, CellValue>>> Tables { get; init; } = new();
}

public record CellValue
{
    public string? RawValue { get; init; }
    public XLDataType DataType { get; init; }
    public T? AsType<T>();
}
```

**Плюсы**:
- Гибкость, не требует объявления моделей
- Быстрая реализация

**Минусы**:
- Нет compile-time type safety
- Неудобно в использовании
- Сложная валидация типов

#### Вариант B: Strongly-typed модель через generics
```csharp
public interface IParsableFromXlsx
{
    static abstract void Configure(XlsxParseProfile profile);
}

public class XlsxParser<TModel> where TModel : IParsableFromXlsx, new()
{
    public TModel Parse(Stream stream);
}
```

**Плюсы**:
- Type safety на уровне компиляции
- IntelliSense поддержка
- Легко тестировать

**Минусы**:
- Требует объявления моделей для каждого типа файлов
- Более сложная реализация

#### Вариант C: Конфигурационная модель (Рекомендуемый)
```csharp
public record XlsxParseResult
{
    public string ProfileName { get; init; } = string.Empty;
    public IReadOnlyList<ParsedField> Fields { get; init; } = new List<ParsedField>();
    public IReadOnlyList<ParsedTable> Tables { get; init; } = new List<ParsedTable>();
    public bool IsSuccess { get; init; }
    public IReadOnlyList<ParseError> Errors { get; init; } = new List<ParseError>();
}

public record ParsedField
{
    public string Name { get; init; } = string.Empty;
    public string? Value { get; init; }
    public XLDataType DataType { get; init; }
    public string? CellAddress { get; init; }
    public T? AsType<T>();
}

public record ParsedTable
{
    public string Name { get; init; } = string.Empty;
    public IReadOnlyList<string> Headers { get; init; } = new List<string>();
    public IReadOnlyList<ParsedTableRow> Rows { get; init; } = new List<ParsedTableRow>();
}

public record ParsedTableRow
{
    public int RowNumber { get; init; }
    public Dictionary<string, ParsedField> Fields { get; init; } = new();
}
```

**Плюсы**:
- Баланс между гибкостью и типизацией
- Единый формат результата для всех профилей
- Методы конвертации типов (`AsType<T>`)
- Легко маппить на domain-модели через extension-методы
- Не требует объявления моделей заранее

**Минусы**:
- Runtime проверка типов (частично компенсируется `AsType<T>`)

### 2. Расширение конфигурации YAML

Добавить секцию `parsing` в существующий профиль:

```yaml
profile: invoice
description: "Входящий счёт от поставщика"
version: "1.0"

validation:
  # ... существующая конфигурация валидации ...

parsing:
  # Маппинг полей на типы данных
  fieldTypes:
    Organization: string
    INN: string
    DocumentDate: date
    TotalAmount: decimal
    Items.Quantity: integer
    Items.Price: decimal
    Items.Sum: decimal

  # Опции парсинга
  options:
    skipEmptyCells: true
    trimStrings: true
    useCulture: "ru-RU"
    dateFormats: ["dd.MM.yyyy", "dd/MM/yyyy"]
    numberStyles: "Number,AllowDecimalPoint"
```

### 3. Архитектурные компоненты

```
src/XlsxValidation/
├── Parsing/
│   ├── IXlsxParser.cs              # Интерфейс парсера
│   ├── XlsxParser.cs               # Основная реализация
│   ├── ParseResult.cs              # Модели результатов (ParsedField, ParsedTable, etc.)
│   ├── TypeConverter.cs            # Конвертация типов данных
│   ├── FieldMapper.cs              # Маппинг полей на модель
│   └── XlsxParseProfile.cs         # Конфигурация парсинга
├── Configuration/
│   ├── XlsxProfileConfig.cs        # Добавить секцию Parsing?
│   └── ...
└── ...
```

### 4. API для использования

```csharp
using Microsoft.Extensions.DependencyInjection;
using XlsxValidation.DependencyInjection;
using XlsxValidation.Parsing;

// Регистрация сервисов
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
    options.EnableParsing = true;  // Новая опция
});

var serviceProvider = services.BuildServiceProvider();
var parserFactory = serviceProvider.GetRequiredService<XlsxParserFactory>();

// Парсинг файла
var parser = parserFactory.CreateForProfile("invoice");
var result = parser.Parse("path/to/invoice.xlsx");

if (result.IsSuccess)
{
    // Доступ к полям
    var organization = result.Fields
        .First(f => f.Name == "Организация")
        .AsString();
    
    var date = result.Fields
        .First(f => f.Name == "Дата документа")
        .AsDateTime();
    
    var total = result.Fields
        .First(f => f.Name == "Итого к оплате")
        .AsDecimal();
    
    // Доступ к таблицам
    var itemsTable = result.Tables.First(t => t.Name == "Позиции");
    foreach (var row in itemsTable.Rows)
    {
        var name = row.Fields["Наименование"].AsString();
        var quantity = row.Fields["Количество"].AsInteger();
        var price = row.Fields["Цена"].AsDecimal();
    }
    
    // Маппинг на domain-модель
    var invoice = result.MapTo<Invoice>();
}
```

### 5. Расширение для маппинга на domain-модели

```csharp
public static class ParseResultExtensions
{
    public static T MapTo<T>(this XlsxParseResult result) where T : new()
    {
        var model = new T();
        var properties = typeof(T).GetProperties();
        
        foreach (var property in properties)
        {
            var fieldAttr = property.GetCustomAttribute<XlsxFieldAttribute>();
            var fieldName = fieldAttr?.Name ?? property.Name;
            
            var field = result.Fields.FirstOrDefault(f => f.Name == fieldName);
            if (field != null)
            {
                var value = field.AsType(property.PropertyType);
                property.SetValue(model, value);
            }
        }
        
        return model;
    }
}

[AttributeUsage(AttributeTargets.Property)]
public class XlsxFieldAttribute : Attribute
{
    public string Name { get; init; }
}

// Пример использования
public class Invoice
{
    [XlsxField(Name = "Организация")]
    public string OrganizationName { get; set; }
    
    [XlsxField(Name = "ИНН")]
    public string INN { get; set; }
    
    [XlsxField(Name = "Дата документа")]
    public DateTime DocumentDate { get; set; }
    
    [XlsxField(Name = "Итого к оплате")]
    public decimal TotalAmount { get; set; }
    
    public List<InvoiceItem> Items { get; set; }
}
```

## Последствия

### Положительные
1. **Переиспользование кода** — система якорей используется и для валидации, и для парсинга
2. **Единая конфигурация** — пользователи описывают структуру файла один раз
3. **Type-safe API** — методы `AsType<T>` обеспечивают безопасную конвертацию
4. **Расширяемость** — возможность добавлять кастомные конвертеры типов
5. **Совместимость** — существующий API валидации не меняется

### Отрицательные
1. **Увеличение сложности** — новые компоненты требуют тестирования и документации
2. **Зависимости** — возможно потребуется добавить System.Text.Json для сериализации
3. **Производительность** — парсинг больших файлов может быть медленным

### Нейтральные
1. **Новые файлы** — потребуется создать ~6-8 новых файлов в папке `Parsing/`
2. **Обучение** — пользователям нужно изучить новый API

## Альтернативы

### Альтернатива 1: Использовать существующие библиотеки
Использовать готовые решения типа **ClosedXML.Report** или **MiniExcel** для парсинга.

**Отклонено**: Не интегрируется с нашей системой якорей и YAML-конфигурацией.

### Альтернатива 2: Только валидация
Оставить проект только для валидации, парсинг — забота пользователей.

**Отклонено**: Ограничивает полезность библиотеки, пользователи просят эту функцию.

### Альтернатива 3: Генерация моделей через T4/Source Generators
Автоматическая генерация C# классов из YAML профиля.

**Отложено**: Может быть добавлено в будущем как enhancement.

## План реализации

### Фаза 1: Базовая инфраструктура
1. Создать папку `Parsing/` с базовыми интерфейсами
2. Реализовать `XlsxParseResult` и связанные модели
3. Добавить `TypeConverter` для конвертации типов

### Фаза 2: Интеграция с конфигурацией
4. Расширить `XlsxProfileConfig` секцией `Parsing`
5. Обновить `YamlProfileLoader` для новой секции
6. Создать `XlsxParserFactory`

### Фаза 3: Реализация парсера
7. Реализовать `XlsxParser` с использованием существующих якорей
8. Добавить маппинг таблиц в `ParsedTable`
9. Реализовать extension-методы для `AsType<T>`

### Фаза 4: Маппинг на модели
10. Создать `XlsxFieldAttribute` и extension-метод `MapTo<T>`
11. Добавить кастомные конвертеры типов

### Фаза 5: Тесты и документация
12. Написать unit-тесты для парсера
13. Обновить README.md примерами парсинга

## Ссылки
- [ClosedXML Documentation](https://closedxml.readthedocs.io/)
- [YamlDotNet Documentation](https://github.com/aaubry/YamlDotNet)
- Existing project structure: `src/XlsxValidation/`
