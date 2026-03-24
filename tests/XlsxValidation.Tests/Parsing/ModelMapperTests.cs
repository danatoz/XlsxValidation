using ClosedXML.Excel;
using XlsxValidation.Parsing;

namespace XlsxValidation.Tests.Parsing;

public class ModelMapperTests
{
    public class MapTo : ModelMapperTests
    {
        [Fact]
        public void Maps_Fields_To_Model_Properties()
        {
            // Arrange
            var result = new XlsxParseResult
            {
                ProfileName = "test",
                Fields = new List<ParsedField>
                {
                    new() { Name = "Name", RawValue = "Test Name", DataType = XLDataType.Text },
                    new() { Name = "Age", RawValue = "30", DataType = XLDataType.Number },
                    new() { Name = "Email", RawValue = "test@example.com", DataType = XLDataType.Text }
                },
                Tables = new List<ParsedTable>()
            };

            // Act
            var model = result.MapTo<TestModel>();

            // Assert
            Assert.Equal("Test Name", model.Name);
            Assert.Equal(30, model.Age);
            Assert.Equal("test@example.com", model.Email);
        }

        [Fact]
        public void Maps_Fields_Using_XlsxField_Attribute()
        {
            // Arrange
            var result = new XlsxParseResult
            {
                ProfileName = "test",
                Fields = new List<ParsedField>
                {
                    new() { Name = "FullName", RawValue = "John Doe", DataType = XLDataType.Text },
                    new() { Name = "YearsOld", RawValue = "25", DataType = XLDataType.Number }
                },
                Tables = new List<ParsedTable>()
            };

            // Act
            var model = result.MapTo<AttributedModel>();

            // Assert
            Assert.Equal("John Doe", model.PersonName);
            Assert.Equal(25, model.Age);
        }

        [Fact]
        public void Maps_Fields_Case_Insensitive()
        {
            // Arrange
            var result = new XlsxParseResult
            {
                ProfileName = "test",
                Fields = new List<ParsedField>
                {
                    new() { Name = "name", RawValue = "Test", DataType = XLDataType.Text },
                    new() { Name = "AGE", RawValue = "40", DataType = XLDataType.Number }
                },
                Tables = new List<ParsedTable>()
            };

            // Act
            var model = result.MapTo<TestModel>();

            // Assert
            Assert.Equal("Test", model.Name);
            Assert.Equal(40, model.Age);
        }

        [Fact]
        public void Handles_Null_Field_Values()
        {
            // Arrange
            var result = new XlsxParseResult
            {
                ProfileName = "test",
                Fields = new List<ParsedField>
                {
                    new() { Name = "Name", RawValue = null, DataType = XLDataType.Text },
                    new() { Name = "Age", RawValue = "30", DataType = XLDataType.Number }
                },
                Tables = new List<ParsedTable>()
            };

            // Act
            var model = result.MapTo<TestModel>();

            // Assert
            Assert.Null(model.Name);
            Assert.Equal(30, model.Age);
        }

        [Fact]
        public void Maps_Existing_Model()
        {
            // Arrange
            var result = new XlsxParseResult
            {
                ProfileName = "test",
                Fields = new List<ParsedField>
                {
                    new() { Name = "Name", RawValue = "Updated Name", DataType = XLDataType.Text }
                },
                Tables = new List<ParsedTable>()
            };

            var model = new TestModel { Age = 20 };

            // Act
            result.MapTo(model);

            // Assert
            Assert.Equal("Updated Name", model.Name);
            Assert.Equal(20, model.Age); // Age should remain unchanged
        }

        [Fact]
        public void Maps_To_Interface_Type()
        {
            // Arrange
            var result = new XlsxParseResult
            {
                ProfileName = "test",
                Fields = new List<ParsedField>
                {
                    new() { Name = "Name", RawValue = "Interface Test", DataType = XLDataType.Text }
                },
                Tables = new List<ParsedTable>()
            };

            // Act
            var model = (ITestInterface)result.MapTo(typeof(ImplementingClass));

            // Assert
            Assert.Equal("Interface Test", model.Name);
        }
    }

    // Test models
    private class TestModel
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public string? Email { get; set; }
    }

    private class AttributedModel
    {
        [XlsxField(Name = "FullName")]
        public string PersonName { get; set; } = string.Empty;

        [XlsxField(Name = "YearsOld")]
        public int Age { get; set; }
    }

    private interface ITestInterface
    {
        string? Name { get; set; }
    }

    private class ImplementingClass : ITestInterface
    {
        public string? Name { get; set; }
    }
}
