# XlsxValidation

Библиотека для валидации XLSX-файлов с конфигурацией через YAML.

## Возможности

- **Декларативная конфигурация** — все правила валидации описываются в YAML-файлах
- **Система якорей** — адресация ячеек через содержимое, смещение, именованные диапазоны или явный адрес
- **Валидация ячеек и таблиц** — поддержка одиночных ячеек и динамических таблиц
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
│   └── XlsxValidation/           # Основная библиотека
├── tests/
│   └── XlsxValidation.Tests/     # Тесты
├── xlsx-profiles/                # YAML-профили валидации
│   ├── _shared.yaml              # Общие наборы правил
│   ├── invoice.yaml
│   ├── salary-report.yaml
│   └── act-of-work.yaml
└── README.md
```

## Запуск тестов

```bash
dotnet test
```

## Требования

- .NET 8.0+
- ClosedXML 0.105.0+
- YamlDotNet 16.3.0+
- Microsoft.Extensions.DependencyInjection.Abstractions 10.0.0+

## Лицензия

MIT
