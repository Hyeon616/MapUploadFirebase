class Program
{
    public static async Task Main(string[] args)
    {
        var logger = new ConsoleLogger();
        var httpClient = new HttpClient();
        var firebaseUrl = "https://enpconventionproject-5ff8f-default-rtdb.firebaseio.com/";
        var firebaseUploader = new FirebaseUploader(httpClient, firebaseUrl, logger);
        var excelProcessor = new ExcuteExcel(firebaseUploader, logger);

        try
        {
            if (args.Length < 1)
            {
                logger.Log("Please provide the upload type (chapter or answer) as an argument.");
                return;
            }

            string uploadType = args[0].ToLower();
            if (uploadType != "chapter" && uploadType != "answer")
            {
                logger.Log("Invalid upload type. Please use either 'chapter' or 'answer'.");
                return;
            }

            string currentDirectory = Directory.GetCurrentDirectory();
            await excelProcessor.ProcessExcelFiles(currentDirectory, uploadType);
        }
        catch (Exception ex)
        {
            logger.Log($"An error occurred: {ex.Message}");
            logger.Log($"Stack Trace: {ex.StackTrace}");
        }

        logger.Log("Processing complete. Press any key to exit...");
        Console.ReadKey();
    }
}
