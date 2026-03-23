using ClosedXML.Excel;

namespace XlsxValidation.Anchors;

/// <summary>
/// Якорь с явным адресом ячейки (fallback для фиксированных форматов)
/// </summary>
public class AddressAnchor : ICellAnchor
{
    private readonly string _address;

    public AddressAnchor(string address)
    {
        _address = address;
    }

    public AnchorResolutionResult Resolve(IXLWorksheet worksheet)
    {
        try
        {
            var cell = worksheet.Cell(_address);
            return AnchorResolutionResult.Success(cell);
        }
        catch (Exception ex)
        {
            return AnchorResolutionResult.Failure(
                $"Некорректный адрес ячейки '{_address}': {ex.Message}");
        }
    }

    public string Description => $"Address: '{_address}'";
}
