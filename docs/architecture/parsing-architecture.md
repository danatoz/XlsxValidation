# Архитектура парсинга XLSX

## Общая схема

```
┌─────────────────────────────────────────────────────────────────┐
│                     YAML Profile (invoice.yaml)                 │
│  profile: invoice                                               │
│  validation: ...                                                │
│  parsing:                                                       │
│    fieldTypes: {...}                                            │
│    options: {...}                                               │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    YamlProfileLoader                            │
│  - Deserializer (YamlDotNet)                                    │
│  - Loads XlsxProfileConfig                                      │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                  XlsxParserFactory                              │
│  - Creates XlsxParser for profile                               │
│  - Caches parsers                                               │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                      XlsxParser                                 │
│  - Uses AnchorFactory for cell location                         │
│  - Reads cells via ClosedXML                                    │
│  - Converts types via TypeConverter                             │
│  - Builds XlsxParseResult                                       │
└─────────────────────────────────────────────────────────────────┘
                              │
              ┌───────────────┴───────────────┐
              │                               │
              ▼                               ▼
┌─────────────────────────┐     ┌─────────────────────────────────┐
│   Single Cells          │     │        Tables                   │
│  - Organization         │     │   - Items[]                     │
│  - INN                  │     │     - No                        │
│  - Date                 │     │     - Name                      │
└─────────────────────────┘     │     - Quantity                  │
                                │     - Price                     │
                                │     - Sum                       │
                                └─────────────────────────────────┘
                                        │
                                        ▼
┌─────────────────────────────────────────────────────────────────┐
│                     XlsxParseResult                             │
│  - ProfileName: "invoice"                                       │
│  - Fields: [ParsedField, ...]                                   │
│  - Tables: [ParsedTable, ...]                                   │
│  - Errors: [ParseError, ...]                                    │
│  - IsSuccess: bool                                              │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│              Extension Methods / Mapper                         │
│  - AsString(), AsDecimal(), AsDateTime()                        │
│  - MapTo<T>() → Domain Model                                    │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                   Domain Model                                  │
│  public class Invoice                                           │
│  {                                                              │
│      public string OrganizationName { get; set; }               │
│      public string INN { get; set; }                            │
│      public DateTime DocumentDate { get; set; }                 │
│      public decimal TotalAmount { get; set; }                   │
│      public List<InvoiceItem> Items { get; set; }               │
│  }                                                              │
└─────────────────────────────────────────────────────────────────┘
```

## Поток данных

```
XLSX File → ClosedXML → Anchor Resolution → Type Conversion → XlsxParseResult → Domain Model
```

## Взаимодействие с валидацией

```
┌──────────────────────────────────────────────────────────────┐
│                    XlsxProfileConfig                         │
│  ┌────────────────────┐    ┌────────────────────────────┐   │
│  │   Validation       │    │      Parsing               │   │
│  │   - Worksheets     │    │   - FieldTypes             │   │
│  │   - Cells          │    │   - Options                │   │
│  │   - Tables         │    │                            │   │
│  │   - Rules          │    │                            │   │
│  └────────────────────┘    └────────────────────────────┘   │
└──────────────────────────────────────────────────────────────┘
           │                                    │
           ▼                                    ▼
┌──────────────────────┐            ┌──────────────────────┐
│  XlsxValidator       │            │   XlsxParser         │
│  - Validate rules    │            │  - Extract values    │
│  - Return errors     │            │  - Convert types     │
│                      │            │  - Build result      │
└──────────────────────┘            └──────────────────────┘
           │                                    │
           ▼                                    ▼
┌──────────────────────┐            ┌──────────────────────┐
│ ValidationReport     │            │  XlsxParseResult     │
│ - IsValid            │            │  - IsSuccess         │
│ - Errors[]           │            │  - Fields[]          │
│ - Warnings[]         │            │  - Tables[]          │
│                      │            │  - Errors[]          │
└──────────────────────┘            └──────────────────────┘
```

## Диаграмма классов

```
┌─────────────────────────────────────────┐
│         XlsxParseResult                 │
├─────────────────────────────────────────┤
│ + ProfileName: string                   │
│ + Fields: IReadOnlyList<ParsedField>    │
│ + Tables: IReadOnlyList<ParsedTable>    │
│ + Errors: IReadOnlyList<ParseError>     │
│ + IsSuccess: bool                       │
├─────────────────────────────────────────┤
│ + AsString(fieldName): string?          │
│ + AsDecimal(fieldName): decimal?        │
│ + AsDateTime(fieldName): DateTime?      │
│ + GetTable(tableName): ParsedTable?     │
│ + MapTo<T>(): T                         │
└─────────────────────────────────────────┘
                    ▲
                    │
        ┌───────────┴───────────┐
        │                       │
┌───────────────┐       ┌───────────────┐
│ ParsedField   │       │ ParsedTable   │
├───────────────┤       ├───────────────┤
│ Name          │       │ Name          │
│ Value         │       │ Headers       │
│ DataType      │       │ Rows          │
│ CellAddress   │       ├───────────────┤
├───────────────┤       │ GetRow(index) │
│ AsString()    │       └───────────────┘
│ AsDecimal()   │
│ AsDateTime()  │
│ AsInteger()   │
│ AsBoolean()   │
│ AsType<T>()   │
└───────────────┘

┌─────────────────────────────────────────┐
│         ParseError                      │
├─────────────────────────────────────────┤
│ + FieldName: string                     │
│ + CellAddress: string?                  │
│ + RowNumber: int?                       │
│ + Message: string                       │
│ + Exception: Exception?                 │
└─────────────────────────────────────────┘
```

## Типы данных и конвертация

```
┌─────────────────────────────────────────────────────────────┐
│                    TypeConverter                            │
├─────────────────────────────────────────────────────────────┤
│  XLDataType → CLR Type                                      │
│  ─────────────────────────────────────────────────────────  │
│  Text       → string                                        │
│  Number     → decimal, double, int, long                    │
│  DateTime   → DateTime, DateOnly, DateTimeOffset            │
│  Boolean    → bool                                          │
│  TimeSpan   → TimeSpan                                      │
│  ─────────────────────────────────────────────────────────  │
│  Custom formats:                                            │
│  - Date: "dd.MM.yyyy", "dd/MM/yyyy"                         │
│  - Number: CultureInfo, NumberStyles                        │
│  - Boolean: "да/нет", "true/false", "1/0"                   │
└─────────────────────────────────────────────────────────────┘
```

## Пример использования

```csharp
// 1. Регистрация сервисов
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
});
var serviceProvider = services.BuildServiceProvider();

// 2. Получение парсера
var parserFactory = serviceProvider.GetRequiredService<XlsxParserFactory>();
var parser = parserFactory.CreateForProfile("invoice");

// 3. Парсинг файла
var result = parser.Parse("invoice-001.xlsx");

// 4. Использование результатов
if (result.IsSuccess)
{
    // Прямой доступ к полям
    var organization = result.Fields
        .First(f => f.Name == "Организация").AsString();
    
    // Доступ через helper-методы
    var total = result.AsDecimal("Итого к оплате");
    var date = result.AsDateTime("Дата документа");
    
    // Работа с таблицами
    var itemsTable = result.GetTable("Позиции");
    foreach (var row in itemsTable.Rows)
    {
        var item = new InvoiceItem
        {
            Name = row.Fields["Наименование"].AsString(),
            Quantity = row.Fields["Количество"].AsInteger(),
            Price = row.Fields["Цена"].AsDecimal(),
            Sum = row.Fields["Сумма"].AsDecimal()
        };
    }
    
    // Маппинг на domain-модель
    var invoice = result.MapTo<Invoice>();
}
else
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"Error: {error.Message}");
    }
}
```
