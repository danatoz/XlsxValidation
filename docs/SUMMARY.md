# Резюме: Развитие проекта XlsxValidation

## Текущее состояние

**XlsxValidation** — библиотека для валидации XLSX файлов с YAML-конфигурацией.

**Основной функционал:**
- ✅ Валидация ячеек и таблиц по правилам
- ✅ Система якорей (content, offset, named-range, address)
- ✅ 13 встроенных правил валидации
- ✅ Поддержка кастомных правил
- ✅ DI-интеграция (Microsoft.Extensions.DependencyInjection)

**Технологический стек:**
- .NET 9.0
- ClosedXML 0.105.0
- YamlDotNet 16.3.0

---

## Предлагаемое развитие

### Цель
Добавить возможность **парсинга** XLSX файлов в структурированные данные с использованием той же YAML-конфигурации.

### Ключевые решения

#### 1. Модель данных — `XlsxParseResult`
```
XlsxParseResult
├── ProfileName: string
├── Fields: ParsedField[]      # Одиночные ячейки
├── Tables: ParsedTable[]      # Таблицы
├── Errors: ParseError[]
└── IsSuccess: bool

ParsedField
├── Name: string
├── Value: string?
├── DataType: XLDataType
├── CellAddress: string?
└── AsType<T>(): T?           # Конвертация типа
```

**Преимущества:**
- Баланс между гибкостью и типизацией
- Методы конвертации (`AsString()`, `AsDecimal()`, `AsDateTime()`)
- Легко маппить на domain-модели

#### 2. Расширение YAML профиля
```yaml
profile: invoice

validation:
  # ... существующие правила валидации ...

parsing:
  fieldTypes:
    Организация: string
    Дата документа: date
    Итого к оплате: decimal
  
  options:
    culture: "ru-RU"
    dateFormats: ["dd.MM.yyyy", "dd/MM/yyyy"]
    trimStrings: true
```

#### 3. API использования
```csharp
// Регистрация
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
    options.EnableParsing = true;
});

// Парсинг
var parser = parserFactory.CreateForProfile("invoice");
var result = parser.Parse("invoice.xlsx");

if (result.IsSuccess)
{
    var organization = result.Fields
        .First(f => f.Name == "Организация").AsString();
    
    var total = result.AsDecimal("Итого к оплате");
    var invoice = result.MapTo<Invoice>();  // Маппинг на модель
}
```

#### 4. Маппинг на domain-модели
```csharp
public class Invoice
{
    [XlsxField(Name = "Организация")]
    public string OrganizationName { get; set; }
    
    [XlsxField(Name = "Итого к оплате")]
    public decimal TotalAmount { get; set; }
    
    public List<InvoiceItem> Items { get; set; }
}

// Использование
var invoice = result.MapTo<Invoice>();
```

---

## Архитектурная схема

```
┌──────────────────┐
│  YAML Profile    │
│  + parsing       │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  XlsxParser      │
│  (использует     │
│   якоря)         │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  XlsxParseResult │
│  - Fields[]      │
│  - Tables[]      │
└────────┬─────────┘
         │
         ▼
┌──────────────────┐
│  MapTo<T>()      │
│  Domain Model    │
└──────────────────┘
```

---

## План реализации (4 недели)

| Неделя | Задачи |
|--------|--------|
| **1** | Модели результатов, TypeConverter, конфигурация |
| **2** | XlsxParser, CellParser, TableParser, extension-методы |
| **3** | Интеграция: DI, Factory, ModelMapper |
| **4** | Тесты (>80% покрытие), документация, примеры |

---

## Созданные документы

| Документ | Описание |
|----------|----------|
| [`docs/adr/001-xlsx-parsing.md`](docs/adr/001-xlsx-parsing.md) | ADR с обоснованием и вариантами |
| [`docs/architecture/parsing-architecture.md`](docs/architecture/parsing-architecture.md) | Диаграммы архитектуры |
| [`docs/implementation-plan.md`](docs/implementation-plan.md) | Детальный план реализации |
| [`xlsx-profiles/invoice-with-parsing.yaml`](xlsx-profiles/invoice-with-parsing.yaml) | Пример профиля с парсингом |
| [`docs/README.md`](docs/README.md) | Индекс документации |

---

## Преимущества решения

✅ **Переиспользование** — система якорей используется и для валидации, и для парсинга  
✅ **Единая конфигурация** — структура файла описывается один раз  
✅ **Type-safe** — методы `AsType<T>` обеспечивают безопасную конвертацию  
✅ **Расширяемость** — кастомные конвертеры, маппинг на модели  
✅ **Обратная совместимость** — существующий API не меняется  

---

## Следующие шаги

1. **Ревью ADR** — обсудить архитектурное решение
2. **Создать issue** — разбить план на задачи в трекере
3. **Начать Фазу 1** — базовая инфраструктура (модели, конвертер)

---

## Контакты

Вопросы и предложения направляйте в репозиторий проекта.
