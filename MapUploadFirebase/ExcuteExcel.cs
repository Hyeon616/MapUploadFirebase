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

    public async Task ProcessExcelFiles(string directory)
    {
        var excelFiles = Directory.GetFiles(directory, "*.xlsm")
                                  .Where(f => !Path.GetFileName(f).StartsWith("~$"));

        foreach (var filePath in excelFiles)
        {
            _logger.Log($"Processing file: {filePath}");
            try
            {
                var chapterData = ProcessExcelFile(filePath);
                var json = JsonConvert.SerializeObject(chapterData);
                await _firebaseUploader.UploadToFirebase(Path.GetFileNameWithoutExtension(filePath), json);
            }
            catch (Exception ex)
            {
                _logger.Log($"An error occurred while processing {filePath}: {ex.Message}");
            }
        }
    }

    private Dictionary<string, object> ProcessExcelFile(string filePath)
    {
        using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var document = SpreadsheetDocument.Open(fs, false);
        var workbookPart = document.WorkbookPart;
        var sheets = workbookPart.Workbook.Descendants<Sheet>();

        var chapterData = new Dictionary<string, object>();
        foreach (var sheet in sheets)
        {
            try
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var stageData = ProcessWorksheet(worksheetPart);
                chapterData[sheet.Name] = stageData;
            }
            catch (Exception ex)
            {
                _logger.Log($"Error processing sheet '{sheet.Name}': {ex.Message}");
            }
        }
        return chapterData;
    }

    private Dictionary<string, object> ProcessWorksheet(WorksheetPart worksheetPart)
    {
        var cellReader = new ReadCell(worksheetPart);
        var mapSize = cellReader.GetIntCellValue("A1", "map size", 1, 7);
        var rotationCount = cellReader.GetIntCellValue("B1", "total puzzle rotation count", 0);

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

        return new Dictionary<string, object>
        {
            {"Size", mapSize},
            {"RotationCount", rotationCount},
            {"Map", map}
        };
    }
}