# XlsxValidation

Библиотека для валидации и парсинга XLSX-файлов с конфигурацией через YAML.

## Возможности

- **Декларативная конфигурация** — все правила валидации и парсинга описываются в YAML-файлах
- **Система якорей** — адресация ячеек через содержимое, смещение, именованные диапазоны или явный адрес
- **Валидация ячеек и таблиц** — поддержка одиночных ячеек и динамических таблиц
- **Парсинг в модели** — извлечение данных в strongly-typed модели
- **Встроенные правила** — 13 готовых правил валидации
- **Кастомные правила** — возможность регистрации собственных правил
- **DI-интеграция** — поддержка Microsoft.Extensions.DependencyInjection

## Установка

```bash
dotnet add package ClosedXML
dotnet add package YamlDotNet
dotnet add package Microsoft.Extensions.DependencyInjection.Abstractions
```

## Быстрый старт

### 1. Создайте профиль валидации

Создайте файл `xlsx-profiles/invoice.yaml`:

```yaml
profile: invoice
description: "Входящий счёт от поставщика"
version: "1.0"

validation:
  worksheets:
    - name: "Данные"

      cells:
        - name: "Организация"
          anchor:
            type: content
            value: "Наименование организации"
          rules:
            - rule: not-empty
            - rule: max-length
              params: { max: 200 }

        - name: "Дата документа"
          anchor:
            type: offset
            base:
              type: content
              value: "Дата составления"
            rowOffset: 0
            colOffset: 1
          rules:
            - rule: not-empty
            - rule: is-date
            - rule: date-not-future

      tables:
        - name: "Позиции"
          headerAnchor:
            type: content
            value: "№"
          stopCondition:
            type: empty-row
          maxRows: 5000
          columns:
            - header: "Наименование"
              rules:
                - rule: not-empty

            - header: "Количество"
              rules:
                - rule: not-empty
                - rule: is-numeric
                - rule: min-value
                  params: { min: 0 }
```

### 2. Используйте в коде

```csharp
using Microsoft.Extensions.DependencyInjection;
using XlsxValidation.DependencyInjection;
using XlsxValidation.Factory;

// Регистрация сервисов
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
});

var serviceProvider = services.BuildServiceProvider();
var validatorFactory = serviceProvider.GetRequiredService<XlsxValidatorFactory>();

// Валидация файла
var validator = validatorFactory.CreateForProfile("invoice");
var report = validator.Validate("path/to/file.xlsx");

if (!report.IsValid)
{
    foreach (var error in report.Errors)
    {
        Console.WriteLine($"Ошибка: {error.FieldName} ({error.CellAddress}) - {error.Message}");
    }
}
```

## Система якорей

Якоря позволяют находить ячейки без привязки к конкретным адресам:

| Тип | Описание | Пример |
|-----|----------|--------|
| `content` | Поиск по содержимому | `type: content, value: "Итого"` |
| `offset` | Смещение от другого якоря | `type: offset, base: {...}, rowOffset: 1, colOffset: 0` |
| `named-range` | Именованный диапазон XLSX | `type: named-range, value: "HeaderCell"` |
| `address` | Явный адрес | `type: address, value: "B3"` |

## Встроенные правила

### Для ячеек и колонок

| Правило | Параметры | Описание |
|---------|-----------|----------|
| `not-empty` | — | Ячейка не должна быть пустой |
| `is-numeric` | — | Значение является числом |
| `is-date` | — | Значение является датой |
| `is-text` | — | Значение является строкой |
| `max-length` | `max: int` | Длина строки не превышает max |
| `min-length` | `min: int` | Длина строки не менее min |
| `min-value` | `min: double` | Числовое значение >= min |
| `max-value` | `max: double` | Числовое значение <= max |
| `matches` | `pattern: string`, `message: string` | Значение соответствует regex |
| `one-of` | `values: [...]` | Значение входит в список допустимых |

### Только для ячеек

| Правило | Описание |
|---------|----------|
| `date-not-future` | Дата не может быть в будущем |
| `date-not-past` | Дата не может быть в прошлом |
| `is-merged` | Ячейка является объединённой |

## Кастомные правила

```csharp
services.AddCustomRule("is-inn", (config, prefix) => cell =>
{
    var value = cell.GetString().Trim();
    var isValid = (value.Length == 10 || value.Length == 12) 
        && value.All(char.IsDigit);
    
    return isValid 
        ? ValidationResult.Ok() 
        : ValidationResult.Error($"{prefix}Некорректный ИНН");
});
```

## Структура проекта

```
xlsxvalidator/
├── src/
│   └── XlsxValidation/
│       ├── Anchors/           # Система якорей
│       ├── Builder/           # Builder для валидатора
│       ├── Configuration/     # Конфигурация и YAML
│       ├── DependencyInjection/
│       ├── Factory/           # Фабрики валидаторов/парсеров
│       ├── Parsing/           # Парсинг XLSX
│       ├── Results/           # Результаты валидации
│       ├── Rules/             # Правила валидации
│       └── Validators/        # Валидаторы
├── tests/
│   └── XlsxValidation.Tests/
├── xlsx-profiles/             # YAML-профили валидации
│   ├── _shared.yaml
│   ├── invoice.yaml
│   ├── salary-report.yaml
│   └── act-of-work.yaml
└── docs/                      # Документация
    ├── adr/                   # Architecture Decision Records
    └── architecture/          # Архитектурные диаграммы
```

## Запуск тестов

```bash
dotnet test
```

## Парсинг XLSX файлов

Библиотека поддерживает парсинг XLSX файлов в структурированные данные с использованием той же YAML-конфигурации.

### 1. Добавьте секцию parsing в профиль

```yaml
profile: invoice
description: "Входящий счёт от поставщика"
version: "1.0"

validation:
  # ... конфигурация валидации ...

parsing:
  # Маппинг полей на типы данных
  fieldTypes:
    Организация: string
    ИНН: string
    Дата документа: date
    Итого к оплате: decimal

  # Опции парсинга
  options:
    skipEmptyCells: true
    trimStrings: true
    culture: "ru-RU"
    dateFormats: ["dd.MM.yyyy", "dd/MM/yyyy"]
```

### 2. Используйте парсер в коде

```csharp
using Microsoft.Extensions.DependencyInjection;
using XlsxValidation.DependencyInjection;
using XlsxValidation.Parsing;

// Регистрация сервисов с включенным парсингом
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
    options.EnableParsing = true;  // Включить парсинг
});

var serviceProvider = services.BuildServiceProvider();
var parserFactory = serviceProvider.GetRequiredService<XlsxParserFactory>();

// Парсинг файла
var parser = parserFactory.CreateForProfile("invoice");
var result = parser.Parse("path/to/file.xlsx");

if (result.IsSuccess)
{
    // Доступ к полям через extension-методы
    var organization = result.Fields
        .First(f => f.Name == "Организация")
        .AsString();

    var inn = result.Fields
        .First(f => f.Name == "ИНН")
        .AsString();

    var date = result.Fields
        .First(f => f.Name == "Дата документа")
        .AsDateTime();

    var total = result.Fields
        .First(f => f.Name == "Итого к оплате")
        .AsDecimal();

    // Доступ к таблицам
    var itemsTable = result.GetTable("Позиции");
    foreach (var row in itemsTable.Rows)
    {
        var name = row.Fields["Наименование"].AsString();
        var quantity = row.Fields["Количество"].AsInteger();
        var price = row.Fields["Цена"].AsDecimal();
    }
}
else
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"Ошибка парсинга: {error.Message}");
    }
}
```

### 3. Маппинг на domain-модель

```csharp
using XlsxValidation.Parsing;

// Определите модель с атрибутами
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

    [XlsxField(Name = "Позиции", Table = "Позиции")]
    public List<InvoiceItem> Items { get; set; }
}

public class InvoiceItem
{
    [XlsxColumn(Header = "Наименование")]
    public string Name { get; set; }

    [XlsxColumn(Header = "Количество")]
    public int Quantity { get; set; }

    [XlsxColumn(Header = "Цена")]
    public decimal Price { get; set; }

    [XlsxColumn(Header = "Сумма")]
    public decimal Sum { get; set; }
}

// Использование маппинга
var invoice = result.MapTo<Invoice>();
Console.WriteLine($"Счёт от {invoice.OrganizationName} на сумму {invoice.TotalAmount}");
```

### Методы конвертации типов

| Метод | Возвращаемый тип | Описание |
|-------|------------------|----------|
| `AsString()` | `string?` | Получить значение как строку |
| `AsInteger()` | `int?` | Получить значение как целое число |
| `AsLong()` | `long?` | Получить значение как long |
| `AsDecimal()` | `decimal?` | Получить значение как decimal |
| `AsDouble()` | `double?` | Получить значение как double |
| `AsDateTime()` | `DateTime?` | Получить значение как дату |
| `AsDateOnly()` | `DateOnly?` | Получить значение как DateOnly |
| `AsBoolean()` | `bool?` | Получить значение как булево |
| `AsTimeSpan()` | `TimeSpan?` | Получить значение как TimeSpan |
| `AsType<T>()` | `T?` | Получить значение как указанный тип |

## Требования

- .NET 8.0+
- ClosedXML 0.105.0+
- YamlDotNet 16.3.0+
- Microsoft.Extensions.DependencyInjection.Abstractions 10.0.0+

## Лицензия

MIT
