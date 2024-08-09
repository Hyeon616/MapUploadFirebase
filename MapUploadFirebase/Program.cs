using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using System.Text;

class Program
{
    private static readonly HttpClient client = new HttpClient();
    private const string FirebaseUrl = "https://enpconventionproject-default-rtdb.firebaseio.com/";

    static async Task Main(string[] args)
    {
        try
        {
            if (args.Length > 0)
            {
                string excelDir = args[0];
                await ProcessExcelFiles(excelDir);
            }
            else
            {
                Console.WriteLine("Please provide the directory path as an argument.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }

    static async Task ProcessExcelFiles(string directory)
    {
        string[] excelFiles = Directory.GetFiles(directory, "*.xlsm")
                                   .Where(f => !Path.GetFileName(f).StartsWith("~$"))
                                   .ToArray();

        foreach (string filePath in excelFiles)
        {
            Console.WriteLine($"Processing file: {filePath}");
            try
            {
                string tempFilePath = Path.GetTempFileName();
                File.Copy(filePath, tempFilePath, true);

                var chapterData = new Dictionary<string, object>();
                using (var fs = new FileStream(tempFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var document = SpreadsheetDocument.Open(fs, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var sheets = workbookPart.Workbook.Descendants<Sheet>();

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
                            Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                            // 이 시트의 처리를 건너뛰고 다음 시트로 계속 진행
                            continue;
                        }
                    }
                }

                File.Delete(tempFilePath);

                string json = JsonConvert.SerializeObject(chapterData);
                await UploadToFirebase(Path.GetFileNameWithoutExtension(filePath), json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing {filePath}: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }
    }

    static Dictionary<string, object> ProcessWorksheet(WorksheetPart worksheetPart)
    {
        string mapSizeValue = GetCellValue(worksheetPart, "A1");
        if (string.IsNullOrEmpty(mapSizeValue))
        {
            throw new InvalidOperationException("Cell A1 is empty or not found");
        }

        if (!int.TryParse(mapSizeValue, out int mapSize))
        {
            throw new FormatException($"Invalid map size in cell A1: {mapSizeValue}. It should be a number.");
        }

        if (mapSize < 1 || mapSize > 7)
        {
            throw new ArgumentOutOfRangeException(nameof(mapSize), $"Invalid map size: {mapSize}. It should be between 1 and 7.");
        }

        var map = new List<List<string>>();

        for (int i = 2; i <= mapSize + 1; i++)
        {
            var row = new List<string>();
            for (int j = 1; j <= mapSize; j++)
            {
                string cellReference = GetCellReference(i, j);
                string cellValue = GetCellValue(worksheetPart, cellReference);
                row.Add(cellValue);
                Console.WriteLine($"Debug: Cell {cellReference} value: '{cellValue}'");
                
            }
            map.Add(row);
        }

        return new Dictionary<string, object>
    {
        {"size", mapSize},
        {"map", map}
    };
    }

    static string GetCellReference(int row, int column)
    {
        string columnLetter = "";
        while (column > 0)
        {
            int modulo = (column - 1) % 26;
            columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
            column = (column - modulo) / 26;
        }
        return columnLetter + row;
    }

    static string GetCellValue(WorksheetPart worksheetPart, string cellReference)
    {
        var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

        if (cell == null)
        {
            Console.WriteLine($"Debug: Cell {cellReference} not found");
            return string.Empty;
        }

        string value = GetActualCellValue(cell, worksheetPart);

        Console.WriteLine($"Debug: Cell {cellReference} value: '{value}'");
        return value;
    }

    static string GetActualCellValue(Cell cell, WorksheetPart worksheetPart)
    {
        if (cell == null)
        {
            return string.Empty;
        }

        // SharedString 처리
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var stringTable = worksheetPart.GetParentParts().OfType<WorkbookPart>().FirstOrDefault()?.SharedStringTablePart?.SharedStringTable;
            if (stringTable != null)
            {
                return stringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
            }
        }

        // 수식 처리
        if (cell.CellFormula != null)
        {
            return $"={cell.CellFormula.Text}";
        }

        // 일반 값 처리
        if (cell.CellValue != null)
        {
            return cell.CellValue.Text;
        }

        return cell.InnerText ?? string.Empty;
    }

    static async Task UploadToFirebase(string chapterName, string jsonData)
    {
        try
        {
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
            var response = await client.PutAsync($"{FirebaseUrl}chapters/{chapterName}.json", content);
            var responseBody = await response.Content.ReadAsStringAsync();
            if (response.IsSuccessStatusCode)
            {
                Console.WriteLine($"Successfully uploaded data for chapter: {chapterName}");
                Console.WriteLine($"Response: {responseBody}");
            }
            else
            {
                Console.WriteLine($"Failed to upload data for chapter: {chapterName}. Status: {response.StatusCode}");
                Console.WriteLine($"Response: {responseBody}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error uploading to Firebase: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }
    }
}