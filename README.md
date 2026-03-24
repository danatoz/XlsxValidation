# XlsxValidation

Library for validating and parsing XLSX files with YAML-based configuration.

## Features

- **Declarative configuration** — all validation and parsing rules are defined in YAML files
- **Anchor system** — cell addressing via content, offset, named ranges, or explicit address
- **Cell and table validation** — support for single cells and dynamic tables
- **Model parsing** — extract data into strongly-typed models
- **Built-in rules** — 13 ready-to-use validation rules
- **Custom rules** — ability to register your own validation rules
- **DI integration** — Microsoft.Extensions.DependencyInjection support

## Installation

```bash
dotnet add package ClosedXML
dotnet add package YamlDotNet
dotnet add package Microsoft.Extensions.DependencyInjection.Abstractions
```

## Quick Start

### 1. Create a validation profile

Create a file `xlsx-profiles/invoice.yaml`:

```yaml
profile: invoice
description: "Incoming invoice from supplier"
version: "1.0"

validation:
  worksheets:
    - name: "Data"

      cells:
        - name: "Organization"
          anchor:
            type: content
            value: "Organization Name"
          rules:
            - rule: not-empty
            - rule: max-length
              params: { max: 200 }

        - name: "Document Date"
          anchor:
            type: offset
            base:
              type: content
              value: "Date of Preparation"
            rowOffset: 0
            colOffset: 1
          rules:
            - rule: not-empty
            - rule: is-date
            - rule: date-not-future

      tables:
        - name: "Items"
          headerAnchor:
            type: content
            value: "No."
          stopCondition:
            type: empty-row
          maxRows: 5000
          columns:
            - header: "Name"
              rules:
                - rule: not-empty

            - header: "Quantity"
              rules:
                - rule: not-empty
                - rule: is-numeric
                - rule: min-value
                  params: { min: 0 }
```

### 2. Use in code

```csharp
using Microsoft.Extensions.DependencyInjection;
using XlsxValidation.DependencyInjection;
using XlsxValidation.Factory;

// Register services
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
});

var serviceProvider = services.BuildServiceProvider();
var validatorFactory = serviceProvider.GetRequiredService<XlsxValidatorFactory>();

// Validate file
var validator = validatorFactory.CreateForProfile("invoice");
var report = validator.Validate("path/to/file.xlsx");

if (!report.IsValid)
{
    foreach (var error in report.Errors)
    {
        Console.WriteLine($"Error: {error.FieldName} ({error.CellAddress}) - {error.Message}");
    }
}
```

## Anchor System

Anchors allow finding cells without binding to specific addresses:

| Type | Description | Example |
|------|-------------|---------|
| `content` | Search by content | `type: content, value: "Total"` |
| `offset` | Offset from another anchor | `type: offset, base: {...}, rowOffset: 1, colOffset: 0` |
| `named-range` | XLSX named range | `type: named-range, value: "HeaderCell"` |
| `address` | Explicit address | `type: address, value: "B3"` |

## Built-in Rules

### For cells and columns

| Rule | Parameters | Description |
|------|------------|-------------|
| `not-empty` | — | Cell must not be empty |
| `is-numeric` | — | Value is a number |
| `is-date` | — | Value is a date |
| `is-text` | — | Value is a string |
| `max-length` | `max: int` | String length does not exceed max |
| `min-length` | `min: int` | String length is at least min |
| `min-value` | `min: double` | Numeric value >= min |
| `max-value` | `max: double` | Numeric value <= max |
| `matches` | `pattern: string`, `message: string` | Value matches regex pattern |
| `one-of` | `values: [...]` | Value is in the allowed list |

### For cells only

| Rule | Description |
|------|-------------|
| `date-not-future` | Date cannot be in the future |
| `date-not-past` | Date cannot be in the past |
| `is-merged` | Cell is merged |

## Custom Rules

```csharp
services.AddCustomRule("is-inn", (config, prefix) => cell =>
{
    var value = cell.GetString().Trim();
    var isValid = (value.Length == 10 || value.Length == 12)
        && value.All(char.IsDigit);

    return isValid
        ? ValidationResult.Ok()
        : ValidationResult.Error($"{prefix}Invalid INN");
});
```

## Project Structure

```
xlsxvalidator/
├── src/
│   └── XlsxValidation/
│       ├── Anchors/           # Anchor system
│       ├── Builder/           # Validator builder
│       ├── Configuration/     # Configuration and YAML
│       ├── DependencyInjection/
│       ├── Factory/           # Validator/Parser factories
│       ├── Parsing/           # XLSX parsing
│       ├── Results/           # Validation results
│       ├── Rules/             # Validation rules
│       └── Validators/        # Validators
├── tests/
│   └── XlsxValidation.Tests/
├── xlsx-profiles/             # YAML validation profiles
│   ├── _shared.yaml
│   ├── invoice.yaml
│   ├── salary-report.yaml
│   └── act-of-work.yaml
└── docs/                      # Documentation
    ├── adr/                   # Architecture Decision Records
    └── architecture/          # Architecture diagrams
```

## Running Tests

```bash
dotnet test
```

## Parsing XLSX Files

The library supports parsing XLSX files into structured data using the same YAML configuration.

### 1. Add parsing section to profile

```yaml
profile: invoice
description: "Incoming invoice from supplier"
version: "1.0"

validation:
  # ... validation configuration ...

parsing:
  # Field type mapping
  fieldTypes:
    Organization: string
    INN: string
    DocumentDate: date
    TotalAmount: decimal

  # Parsing options
  options:
    skipEmptyCells: true
    trimStrings: true
    culture: "ru-RU"
    dateFormats: ["dd.MM.yyyy", "dd/MM/yyyy"]
```

### 2. Use parser in code

```csharp
using Microsoft.Extensions.DependencyInjection;
using XlsxValidation.DependencyInjection;
using XlsxValidation.Parsing;

// Register services with parsing enabled
var services = new ServiceCollection();
services.AddXlsxValidation(options =>
{
    options.ProfilesDirectory = "xlsx-profiles";
    options.EnableParsing = true;  // Enable parsing
});

var serviceProvider = services.BuildServiceProvider();
var parserFactory = serviceProvider.GetRequiredService<XlsxParserFactory>();

// Parse file
var parser = parserFactory.CreateForProfile("invoice");
var result = parser.Parse("path/to/file.xlsx");

if (result.IsSuccess)
{
    // Access fields via extension methods
    var organization = result.Fields
        .First(f => f.Name == "Organization")
        .AsString();

    var inn = result.Fields
        .First(f => f.Name == "INN")
        .AsString();

    var date = result.Fields
        .First(f => f.Name == "DocumentDate")
        .AsDateTime();

    var total = result.Fields
        .First(f => f.Name == "TotalAmount")
        .AsDecimal();

    // Access tables
    var itemsTable = result.GetTable("Items");
    foreach (var row in itemsTable.Rows)
    {
        var name = row.Fields["Name"].AsString();
        var quantity = row.Fields["Quantity"].AsInteger();
        var price = row.Fields["Price"].AsDecimal();
    }
}
else
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"Parse error: {error.Message}");
    }
}
```

### 3. Mapping to domain model

```csharp
using XlsxValidation.Parsing;

// Define model with attributes
public class Invoice
{
    [XlsxField(Name = "Organization")]
    public string OrganizationName { get; set; }

    [XlsxField(Name = "INN")]
    public string INN { get; set; }

    [XlsxField(Name = "DocumentDate")]
    public DateTime DocumentDate { get; set; }

    [XlsxField(Name = "TotalAmount")]
    public decimal TotalAmount { get; set; }

    [XlsxField(Name = "Items", Table = "Items")]
    public List<InvoiceItem> Items { get; set; }
}

public class InvoiceItem
{
    [XlsxColumn(Header = "Name")]
    public string Name { get; set; }

    [XlsxColumn(Header = "Quantity")]
    public int Quantity { get; set; }

    [XlsxColumn(Header = "Price")]
    public decimal Price { get; set; }

    [XlsxColumn(Header = "Sum")]
    public decimal Sum { get; set; }
}

// Use mapping
var invoice = result.MapTo<Invoice>();
Console.WriteLine($"Invoice from {invoice.OrganizationName} for {invoice.TotalAmount}");
```

### Type Conversion Methods

| Method | Return Type | Description |
|--------|-------------|-------------|
| `AsString()` | `string?` | Get value as string |
| `AsInteger()` | `int?` | Get value as integer |
| `AsLong()` | `long?` | Get value as long |
| `AsDecimal()` | `decimal?` | Get value as decimal |
| `AsDouble()` | `double?` | Get value as double |
| `AsDateTime()` | `DateTime?` | Get value as date |
| `AsDateOnly()` | `DateOnly?` | Get value as DateOnly |
| `AsBoolean()` | `bool?` | Get value as boolean |
| `AsTimeSpan()` | `TimeSpan?` | Get value as TimeSpan |
| `AsType<T>()` | `T?` | Get value as specified type |

## Requirements

- .NET 8.0+
- ClosedXML 0.105.0+
- YamlDotNet 16.3.0+
- Microsoft.Extensions.DependencyInjection.Abstractions 10.0.0+

## License

MIT
