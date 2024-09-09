using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;

public class ExcuteExcel
{
    private readonly IFirebaseUploader _firebaseUploader;
    private readonly ILogger _logger;

    public ExcuteExcel(IFirebaseUploader firebaseUploader, ILogger logger)
    {
        _firebaseUploader = firebaseUploader ?? throw new ArgumentNullException(nameof(firebaseUploader));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public async Task ProcessExcelFiles(string directory, string uploadType)
    {
        var excelFiles = Directory.GetFiles(directory, "*.xlsm")
                                  .Where(f => !Path.GetFileName(f).StartsWith("~$"));

        foreach (var filePath in excelFiles)
        {
            _logger.Log($"Processing file: {filePath}");
            try
            {
                var data = ProcessExcelFile(filePath, uploadType);
                var json = JsonConvert.SerializeObject(data);

                if (uploadType.ToLower() == "chapter")
                {
                    await _firebaseUploader.UploadToFirebaseChapter(Path.GetFileNameWithoutExtension(filePath), json);
                }
                else if (uploadType.ToLower() == "answer")
                {
                    await _firebaseUploader.UploadToFirebaseAnswers(Path.GetFileNameWithoutExtension(filePath), json);
                }
            }
            catch (Exception ex)
            {
                _logger.Log($"An error occurred while processing {filePath}: {ex.Message}");
            }
        }
    }

    private Dictionary<string, object> ProcessExcelFile(string filePath, string uploadType)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var document = SpreadsheetDocument.Open(fs, false);
        var workbookPart = document.WorkbookPart;
        var sheets = workbookPart.Workbook.Descendants<Sheet>();

        var data = new Dictionary<string, object>();
        foreach (var sheet in sheets)
        {
            try
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var stageData = ProcessWorksheet(worksheetPart, uploadType);
                data[sheet.Name] = stageData;
            }
            catch (Exception ex)
            {
                _logger.Log($"Error processing sheet '{sheet.Name}': {ex.Message}");
            }
        }
        return data;
    }

    private Dictionary<string, object> ProcessWorksheet(WorksheetPart worksheetPart, string uploadType)
    {
        var cellReader = new ReadCell(worksheetPart);
        var mapSize = cellReader.GetIntCellValue("A1", "map size", 1, 7);

        var map = new List<List<string>>();
        for (int row = 2; row <= mapSize + 1; row++)
        {
            var rowData = new List<string>();
            for (int col = 1; col <= mapSize; col++)
            {
                string cellReference = $"{(char)('A' + col - 1)}{row}";
                string cellValue = cellReader.GetCellValue(cellReference);
                rowData.Add(cellValue);
            }
            map.Add(rowData);
        }

        var result = new Dictionary<string, object>
        {
            {"Map", map}
        };

        if (uploadType.ToLower() == "chapter")
        {
            var rotationCount = cellReader.GetIntCellValue("B1", "total puzzle rotation count", 0);

            var blockedCellsString = cellReader.GetCellValue("C1");
            var blockedCells = new List<string>();

            if (!string.IsNullOrEmpty(blockedCellsString))
            {
                var cellReferences = blockedCellsString.Split(',', StringSplitOptions.RemoveEmptyEntries);
                foreach (var cellRef in cellReferences)
                {
                    var trimmedRef = cellRef.Trim();
                    if (char.IsLetter(trimmedRef[0]) && char.IsDigit(trimmedRef[1]))
                    {
                        int col = trimmedRef[0] - 'A';
                        int row = int.Parse(trimmedRef.Substring(1)) - 2; // Subtract 2 because map starts from A2
                        blockedCells.Add($"{row}_{col}");
                    }
                }
            }

            var sequenceString = cellReader.GetCellValue("D1");
            var sequence = sequenceString.Split(',', StringSplitOptions.RemoveEmptyEntries)
                                         .Select(s => s.Trim())
                                         .ToList();

            result["Size"] = mapSize;
            result["RotationCount"] = rotationCount;
            result["Block"] = blockedCells;
            result["Sequence"] = sequence;
        }

        return result;
    }
}