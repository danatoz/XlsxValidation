using ClosedXML.Excel;

namespace XlsxValidation.Anchors;

/// <summary>
/// Якорь, использующий именованный диапазон XLSX
/// </summary>
public class NamedRangeAnchor : ICellAnchor
{
    private readonly string _rangeName;

    public NamedRangeAnchor(string rangeName)
    {
        _rangeName = rangeName;
    }

    public AnchorResolutionResult Resolve(IXLWorksheet worksheet)
    {
        // Поиск в именованных диапазонах книги
        var workbook = worksheet.Workbook;
        
        // Сначала ищем в именованных диапазонах листа
        var namedRange = worksheet.NamedRange(_rangeName);
        
        // Если не найдено, ищем в именованных диапазонах книги
        if (namedRange == null || !namedRange.Ranges.Any())
            namedRange = workbook.NamedRange(_rangeName);

        if (namedRange == null || !namedRange.Ranges.Any())
            return AnchorResolutionResult.Failure(
                $"Именованный диапазон '{_rangeName}' не найден");

        // Получаем первую ячейку из диапазона
        var range = namedRange.Ranges.First();
        var cell = range.FirstCell();
        
        return AnchorResolutionResult.Success(cell);
    }

    public string Description => $"Named Range: '{_rangeName}'";
}
