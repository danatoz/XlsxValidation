using FluentAssertions;
using XlsxValidation.Results;

namespace XlsxValidation.Tests.Results;

/// <summary>
/// Тесты для ValidationReport и связанных типов
/// </summary>
public class ValidationReportTests
{
    [Fact]
    public void ValidationReport_Empty_IsValid()
    {
        // Arrange
        var report = ValidationReport.Empty("test-profile");

        // Assert
        report.IsValid.Should().BeTrue();
        report.ProfileName.Should().Be("test-profile");
        report.Errors.Should().BeEmpty();
    }

    [Fact]
    public void ValidationReport_WithErrors_IsNotValid()
    {
        // Arrange
        var errors = new List<ValidationError>
        {
            new ValidationError
            {
                FieldName = "Field1",
                CellAddress = "A1",
                RuleId = "not-empty",
                Message = "Ячейка не должна быть пустой"
            }
        };

        // Act
        var report = ValidationReport.WithErrors("test-profile", errors);

        // Assert
        report.IsValid.Should().BeFalse();
        report.Errors.Should().HaveCount(1);
    }

    [Fact]
    public void ValidationReport_Success_IsValid()
    {
        // Arrange & Act
        var report = ValidationReport.Success("test-profile");

        // Assert
        report.IsValid.Should().BeTrue();
        report.Errors.Should().BeEmpty();
    }

    [Fact]
    public void ValidationError_ToString_ReturnsFormattedString()
    {
        // Arrange
        var error = new ValidationError
        {
            FieldName = "ИНН",
            CellAddress = "B2",
            RuleId = "matches",
            Message = "ИНН должен содержать 10 или 12 цифр"
        };

        // Act
        var str = error.ToString();

        // Assert
        str.Should().Contain("ИНН");
        str.Should().Contain("B2");
        str.Should().Contain("matches");
    }

    [Fact]
    public void ValidationError_WithWorksheetName_IncludesInString()
    {
        // Arrange
        var error = new ValidationError
        {
            FieldName = "Поле",
            CellAddress = "A1",
            RuleId = "not-empty",
            Message = "Ошибка",
            WorksheetName = "Данные"
        };

        // Act
        var str = error.ToString();

        // Assert
        str.Should().Contain("Данные");
        str.Should().Contain("A1");
    }

    [Fact]
    public void ValidationError_WithoutCellAddress_UsesRowNumber()
    {
        // Arrange
        var error = new ValidationError
        {
            FieldName = "Поле",
            RowNumber = 5,
            RuleId = "not-empty",
            Message = "Ошибка"
        };

        // Act
        var str = error.ToString();

        // Assert
        str.Should().Contain("строка 5");
    }

    [Fact]
    public void ValidationReport_ReadOnlyErrors_CannotModify()
    {
        // Arrange
        var report = ValidationReport.Empty("test");

        // Assert
        report.Errors.Should().BeOfType<List<ValidationError>>()
            .Which.AsReadOnly().Should().NotBeNull();
    }

    [Fact]
    public void ValidationReport_MultipleErrors_AllIncluded()
    {
        // Arrange
        var errors = new List<ValidationError>
        {
            new ValidationError { FieldName = "Field1", RuleId = "rule1", Message = "Error1" },
            new ValidationError { FieldName = "Field2", RuleId = "rule2", Message = "Error2" },
            new ValidationError { FieldName = "Field3", RuleId = "rule3", Message = "Error3" }
        };

        // Act
        var report = ValidationReport.WithErrors("test-profile", errors);

        // Assert
        report.Errors.Should().HaveCount(3);
        report.Errors.Select(e => e.FieldName).Should().Contain("Field1", "Field2", "Field3");
    }

    [Fact]
    public void ValidationWarning_HasAllProperties()
    {
        // Arrange
        var warning = new ValidationWarning
        {
            FieldName = "Поле",
            CellAddress = "A1",
            RowNumber = 10,
            RuleId = "custom-rule",
            Message = "Предупреждение",
            WorksheetName = "Лист1"
        };

        // Assert
        warning.FieldName.Should().Be("Поле");
        warning.CellAddress.Should().Be("A1");
        warning.RowNumber.Should().Be(10);
        warning.RuleId.Should().Be("custom-rule");
        warning.Message.Should().Be("Предупреждение");
        warning.WorksheetName.Should().Be("Лист1");
    }
}
