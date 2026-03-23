using FluentAssertions;
using XlsxValidation.Configuration;

namespace XlsxValidation.Tests.Configuration;

/// <summary>
/// Тесты для YamlProfileLoader
/// </summary>
public class YamlProfileLoaderTests : IDisposable
{
    private readonly string _testDirectory;
    private readonly YamlProfileLoader _loader;

    public YamlProfileLoaderTests()
    {
        _testDirectory = Path.Combine(Path.GetTempPath(), $"xlsx_test_{Guid.NewGuid()}");
        Directory.CreateDirectory(_testDirectory);
        _loader = new YamlProfileLoader();
    }

    [Fact]
    public void LoadFile_ValidYaml_DeserializesSuccessfully()
    {
        // Arrange
        var yaml = @"
profile: test-profile
description: ""Тестовый профиль""
version: ""1.0""

validation:
  worksheets:
    - name: ""Данные""
      cells:
        - name: ""Поле1""
          anchor:
            type: content
            value: ""Заголовок""
          rules:
            - rule: not-empty
";
        var filePath = Path.Combine(_testDirectory, "test.yaml");
        File.WriteAllText(filePath, yaml);

        // Act
        var config = _loader.LoadFile(filePath);

        // Assert
        config.Profile.Should().Be("test-profile");
        config.Description.Should().Be("Тестовый профиль");
        config.Validation.Worksheets.Should().HaveCount(1);
        config.Validation.Worksheets[0].Cells.Should().HaveCount(1);
    }

    [Fact]
    public void LoadFile_NonExistingFile_ThrowsException()
    {
        // Arrange
        var filePath = Path.Combine(_testDirectory, "non-existing.yaml");

        // Act & Assert
        _loader.Invoking(l => l.LoadFile(filePath))
            .Should().Throw<FileNotFoundException>();
    }

    [Fact]
    public void LoadFile_InvalidYaml_ThrowsProfileLoadException()
    {
        // Arrange
        var yaml = "invalid: yaml: content: [";
        var filePath = Path.Combine(_testDirectory, "invalid.yaml");
        File.WriteAllText(filePath, yaml);

        // Act & Assert
        _loader.Invoking(l => l.LoadFile(filePath))
            .Should().Throw<ProfileLoadException>();
    }

    [Fact]
    public void LoadDirectory_LoadsAllNonSharedProfiles()
    {
        // Arrange
        var profile1 = @"
profile: profile-1
validation:
  worksheets: []
";
        var profile2 = @"
profile: profile-2
validation:
  worksheets: []
";
        var shared = @"
_rules:
  test: &test
    - rule: not-empty
";

        File.WriteAllText(Path.Combine(_testDirectory, "profile1.yaml"), profile1);
        File.WriteAllText(Path.Combine(_testDirectory, "profile2.yml"), profile2);
        File.WriteAllText(Path.Combine(_testDirectory, "_shared.yaml"), shared);

        // Act
        var profiles = _loader.LoadDirectory(_testDirectory);

        // Assert
        profiles.Should().HaveCount(2);
        profiles.ContainsKey("profile-1").Should().BeTrue();
        profiles.ContainsKey("profile-2").Should().BeTrue();
    }

    [Fact]
    public void LoadDirectory_EmptyDirectory_ReturnsEmptyDictionary()
    {
        // Act
        var profiles = _loader.LoadDirectory(_testDirectory);

        // Assert
        profiles.Should().BeEmpty();
    }

    [Fact]
    public void LoadDirectory_NonExistingDirectory_ThrowsDirectoryNotFoundException()
    {
        // Arrange
        var nonExistingDir = Path.Combine(_testDirectory, "non-existing");

        // Act & Assert
        _loader.Invoking(l => l.LoadDirectory(nonExistingDir))
            .Should().Throw<DirectoryNotFoundException>();
    }

    [Fact]
    public void LoadYaml_ValidYaml_DeserializesSuccessfully()
    {
        // Arrange
        var yaml = @"
profile: inline-profile
description: ""Inline YAML""
validation:
  worksheets:
    - name: ""Sheet1""
      tables:
        - name: ""Table1""
          headerAnchor:
            type: content
            value: ""№""
          columns:
            - header: ""Колонка1""
              rules:
                - rule: not-empty
";

        // Act
        var config = _loader.LoadYaml(yaml);

        // Assert
        config.Profile.Should().Be("inline-profile");
        config.Validation.Worksheets[0].Tables.Should().HaveCount(1);
    }

    [Fact]
    public void LoadYaml_InvalidYaml_ThrowsProfileLoadException()
    {
        // Arrange
        var yaml = "invalid: yaml: [";

        // Act & Assert
        _loader.Invoking(l => l.LoadYaml(yaml))
            .Should().Throw<ProfileLoadException>();
    }

    [Fact]
    public void LoadDirectory_WithYmlExtension_LoadsSuccessfully()
    {
        // Arrange
        var yaml = @"
profile: yml-profile
validation:
  worksheets: []
";
        File.WriteAllText(Path.Combine(_testDirectory, "test.yml"), yaml);

        // Act
        var profiles = _loader.LoadDirectory(_testDirectory);

        // Assert
        profiles.Should().HaveCount(1);
        profiles.ContainsKey("yml-profile").Should().BeTrue();
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDirectory))
                Directory.Delete(_testDirectory, true);
        }
        catch
        {
            // Игнорируем ошибки очистки
        }
    }
}
