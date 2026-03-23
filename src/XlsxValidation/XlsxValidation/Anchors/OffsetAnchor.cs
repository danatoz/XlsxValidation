using ClosedXML.Excel;

namespace XlsxValidation.Anchors;

/// <summary>
/// Якорь, смещающийся относительно другого якоря
/// </summary>
public class OffsetAnchor : ICellAnchor
{
    private readonly ICellAnchor _baseAnchor;
    private readonly int _rowOffset;
    private readonly int _colOffset;

    public OffsetAnchor(ICellAnchor baseAnchor, int rowOffset, int colOffset)
    {
        _baseAnchor = baseAnchor;
        _rowOffset = rowOffset;
        _colOffset = colOffset;
    }

    public AnchorResolutionResult Resolve(IXLWorksheet worksheet)
    {
        var baseResult = _baseAnchor.Resolve(worksheet);
        
        if (!baseResult.IsSuccess || baseResult.Cell == null)
            return AnchorResolutionResult.Failure(
                $"Базовый якорь не найден: {baseResult.ErrorMessage}");

        var baseCell = baseResult.Cell;
        var targetRow = baseCell.Address.RowNumber + _rowOffset;
        var targetColumn = baseCell.Address.ColumnNumber + _colOffset;

        // Проверка границ листа
        if (targetRow < 1 || targetRow > XLHelper.MaxRowNumber)
            return AnchorResolutionResult.Failure(
                $"Смещение выходит за границы листа по строке: {targetRow}");

        if (targetColumn < 1 || targetColumn > XLHelper.MaxColumnNumber)
            return AnchorResolutionResult.Failure(
                $"Смещение выходит за границы листа по колонке: {targetColumn}");

        var targetCell = worksheet.Cell(targetRow, targetColumn);
        return AnchorResolutionResult.Success(targetCell);
    }

    public string Description => 
        $"Offset от {_baseAnchor.Description}: row={_rowOffset:+#;-#;0}, col={_colOffset:+#;-#;0}";
}
