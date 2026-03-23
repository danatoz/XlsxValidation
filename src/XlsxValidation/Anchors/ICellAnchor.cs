using ClosedXML.Excel;

namespace XlsxValidation.Anchors;

/// <summary>
/// Результат разрешения якоря
/// </summary>
public record AnchorResolutionResult
{
    /// <summary>
    /// Найденная ячейка (null если якорь не найден)
    /// </summary>
    public IXLCell? Cell { get; init; }

    /// <summary>
    /// Сообщение об ошибке (если якорь не найден)
    /// </summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Был ли якорь успешно разрешён
    /// </summary>
    public bool IsSuccess => Cell != null && ErrorMessage == null;

    public static AnchorResolutionResult Success(IXLCell cell) => new() { Cell = cell };
    public static AnchorResolutionResult Failure(string message) => new() { ErrorMessage = message };
}

/// <summary>
/// Интерфейс якоря для поиска ячейки на листе
/// </summary>
public interface ICellAnchor
{
    /// <summary>
    /// Разрешить якорь в конкретную ячейку
    /// </summary>
    /// <param name="worksheet">Лист для поиска</param>
    /// <returns>Результат разрешения якоря</returns>
    AnchorResolutionResult Resolve(IXLWorksheet worksheet);

    /// <summary>
    /// Логическое описание якоря (для отладки и сообщений об ошибках)
    /// </summary>
    string Description { get; }
}
