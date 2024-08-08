using System;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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
        if (worksheetPart == null || worksheetPart.Worksheet == null)
        {
            throw new ArgumentNullException(nameof(worksheetPart), "WorksheetPart or Worksheet is null");
        }

        var cells = worksheetPart.Worksheet.Descendants<Cell>().ToList();
        if (cells == null || !cells.Any())
        {
            throw new InvalidOperationException("No cells found in the worksheet");
        }

        var sharedStringPart = worksheetPart.GetParentParts().OfType<SharedStringTablePart>().FirstOrDefault();

        string mapSizeValue = GetCellValue(cells, "A1", sharedStringPart);
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
        var validValues = new HashSet<string>
        {
            "C1", "C2", "C3", "H1", "H2", "H3", "O1", "O2",
            "HAM", "CAT", "DOG", "OTT", "PEN", "RAC", "CLO", "BLU", "SEA", "TUR", "FOX", "KAN", "ARM", "OST", "LIO", "WAR", "MEE", "CAP", "MON", "JAG", "PAR", "TOU", "RAB", "POL", "ARC", "SEL",
            "FHAM", "FCAT", "FDOG", "FOTT", "FPEN", "FRAC", "FCLO", "FBLU", "FSEA", "FTUR", "FFOX", "FKAN", "FARM", "FOST", "FLIO", "FWAR", "FMEE", "FCAP", "FMON", "FJAG", "FPAR", "FTOU", "FRAB", "FPOL", "FARC", "FSEL"
        };

        for (int i = 2; i <= mapSize + 1; i++)
        {
            var row = new List<string>();
            for (int j = 1; j <= mapSize; j++)
            {
                string cellReference = GetCellReference(i, j);
                string cellValue = GetCellValue(cells, cellReference, sharedStringPart);

                if (string.IsNullOrEmpty(cellValue))
                {
                    throw new Exception($"Empty cell at [{i}, {j}]");
                }
                if (!validValues.Contains(cellValue))
                {
                    throw new Exception($"Invalid value '{cellValue}' at cell [{i}, {j}]");
                }
                row.Add(cellValue);
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

    static string GetCellValue(IEnumerable<Cell> cells, string cellReference, SharedStringTablePart sharedStringPart)
    {
        var cell = cells.FirstOrDefault(c => c.CellReference == cellReference);
        if (cell == null)
        {
            return string.Empty;
        }

        string value = cell.InnerText;
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            if (sharedStringPart?.SharedStringTable != null)
            {
                return sharedStringPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
        }
        return value;
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