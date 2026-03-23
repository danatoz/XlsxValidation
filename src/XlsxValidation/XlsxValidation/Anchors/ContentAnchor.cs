using ClosedXML.Excel;

namespace XlsxValidation.Anchors;

/// <summary>
/// Якорь, ищущий ячейку по содержимому
/// </summary>
public class ContentAnchor : ICellAnchor
{
    private readonly string _searchValue;
    private readonly bool _exactMatch;
    private readonly int? _occurrence;

    public ContentAnchor(string searchValue, bool exactMatch = false, int? occurrence = null)
    {
        _searchValue = searchValue;
        _exactMatch = exactMatch;
        _occurrence = occurrence;
    }

    public AnchorResolutionResult Resolve(IXLWorksheet worksheet)
    {
        var cells = worksheet.CellsUsed();
        var matches = new List<IXLCell>();

        foreach (var cell in cells)
        {
            var cellValue = cell.GetValue<string>();
            if (string.IsNullOrEmpty(cellValue))
                continue;

            bool isMatch = _exactMatch
                ? cellValue.Trim() == _searchValue.Trim()
                : cellValue.Contains(_searchValue, StringComparison.OrdinalIgnoreCase);

            if (isMatch)
                matches.Add(cell);
        }

        if (matches.Count == 0)
            return AnchorResolutionResult.Failure(
                $"Ячейка со значением '{_searchValue}' не найдена");

        // Если указана конкретная вхождаемость (1-based)
        if (_occurrence.HasValue)
        {
            var index = _occurrence.Value - 1;
            if (index < 0 || index >= matches.Count)
                return AnchorResolutionResult.Failure(
                    $"Ячейка со значением '{_searchValue}' (вхождение #{_occurrence}) не найдена. Найдено вхождений: {matches.Count}");
            
            return AnchorResolutionResult.Success(matches[index]);
        }

        // Возвращаем первое найденное совпадение
        return AnchorResolutionResult.Success(matches.First());
    }

    public string Description => $"Content: '{_searchValue}'{( _exactMatch ? " (точное)" : "")}";
}
