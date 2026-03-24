# Документация XlsxValidation

## Архитектурные решения (ADR)

- **[ADR-001](adr/001-xlsx-parsing.md)** — Добавление парсинга XLSX файлов с YAML-конфигурацией

## Архитектура

- **[Схема парсинга](architecture/parsing-architecture.md)** — Диаграммы и описание архитектуры парсинга

## Планы

- **[План реализации](implementation-plan.md)** — Пошаговый план реализации парсинга

## Профили

Примеры YAML профилей находятся в директории `xlsx-profiles/`:

- `invoice.yaml` — Базовый профиль валидации для счёта
- `invoice-with-parsing.yaml` — Расширенный профиль с секцией парсинга
- `salary-report.yaml` — Профиль для отчёта по зарплате
- `act-of-work.yaml` — Профиль для акта выполненных работ

## Быстрый старт

### Валидация

```csharp
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
});

var validatorFactory = services.BuildServiceProvider()
    .GetRequiredService<XlsxValidatorFactory>();

var validator = validatorFactory.CreateForProfile("invoice");
var report = validator.Validate("file.xlsx");
```

### Парсинг (планируется)

```csharp
var parserFactory = services.BuildServiceProvider()
    .GetRequiredService<XlsxParserFactory>();

var parser = parserFactory.CreateForProfile("invoice");
var result = parser.Parse("file.xlsx");

var organization = result.Fields
    .First(f => f.Name == "Организация").AsString();
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
│       ├── Parsing/           # Парсинг (в разработке)
│       ├── Results/           # Результаты валидации
│       ├── Rules/             # Правила валидации
│       └── Validators/        # Валидаторы
├── tests/
│   └── XlsxValidation.Tests/
├── xlsx-profiles/             # YAML профили
└── docs/                      # Документация
    ├── adr/
    ├── architecture/
    └── ...
```

## Встроенные правила валидации

| Правило | Описание |
|---------|----------|
| `not-empty` | Ячейка не должна быть пустой |
| `is-numeric` | Значение является числом |
| `is-date` | Значение является датой |
| `is-text` | Значение является строкой |
| `max-length` | Длина строки не превышает max |
| `min-length` | Длина строки не менее min |
| `min-value` | Числовое значение >= min |
| `max-value` | Числовое значение <= max |
| `matches` | Значение соответствует regex |
| `one-of` | Значение входит в список допустимых |
| `date-not-future` | Дата не может быть в будущем |
| `date-not-past` | Дата не может быть в прошлом |
| `is-merged` | Ячейка является объединённой |

## Система якорей

| Тип | Описание | Пример |
|-----|----------|--------|
| `content` | Поиск по содержимому | `type: content, value: "Итого"` |
| `offset` | Смещение от другого якоря | `type: offset, base: {...}, rowOffset: 1` |
| `named-range` | Именованный диапазон XLSX | `type: named-range, value: "HeaderCell"` |
| `address` | Явный адрес | `type: address, value: "B3"` |

## Лицензия

MIT
