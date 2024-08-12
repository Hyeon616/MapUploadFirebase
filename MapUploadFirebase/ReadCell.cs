using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

public class ReadCell
{
    private readonly WorksheetPart _worksheetPart;

    public ReadCell(WorksheetPart worksheetPart)
    {
        _worksheetPart = worksheetPart;
    }

    public int GetIntCellValue(string cellReference, string valueName, int minValue, int? maxValue = null)
    {
        string value = GetCellValue(cellReference);
        if (string.IsNullOrEmpty(value))
        {
            Console.WriteLine($"Warning: Cell {cellReference} is empty. Using default value 0.");
            return 0;
        }

        if (!int.TryParse(value, out int result))
        {
            throw new FormatException($"Invalid {valueName} in cell {cellReference}: {value}. It should be a number.");
        }

        if (result < minValue || (maxValue.HasValue && result > maxValue.Value))
        {
            throw new ArgumentOutOfRangeException(nameof(result), $"Invalid {valueName}: {result}. It should be between {minValue} and {maxValue ?? int.MaxValue}.");
        }

        return result;
    }

    public string GetCellValue(string cellReference)
    {
        var cell = _worksheetPart.Worksheet.Descendants<Cell>()
            .FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.OrdinalIgnoreCase));

        if (cell == null)
        {
            return string.Empty;
        }

        var value = GetActualCellValue(cell);
        return value;
    }

    private string GetActualCellValue(Cell cell)
    {
        if (cell.DataType != null)
        {
            if (cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = _worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault()?.SharedStringTablePart?.SharedStringTable;
                if (stringTable != null)
                {
                    return stringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                }
            }
            else if (cell.DataType.Value == CellValues.Boolean)
            {
                return cell.InnerText == "1" ? "TRUE" : "FALSE";
            }
        }

        if (cell.CellValue != null)
        {
            return cell.CellValue.Text;
        }

        return cell.InnerText ?? string.Empty;
    }
}