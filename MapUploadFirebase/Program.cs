class Program
{
    public static async Task Main(string[] args)
    {
        var logger = new ConsoleLogger();
        logger.Log("Program started.");
        logger.Log($"Current Working Directory: {Directory.GetCurrentDirectory()}");

        try
        {
            string directoryPath = Directory.GetCurrentDirectory();
            logger.Log($"Using Directory: {directoryPath}");

            string uploadType;
            do
            {
                Console.Write("Enter upload type (chapter or answer): ");
                uploadType = Console.ReadLine().ToLower();
            } while (uploadType != "chapter" && uploadType != "answer");

            logger.Log($"Upload Type: {uploadType}");

            // 디렉토리 내용 출력
            logger.Log("Directory contents:");
            foreach (string file in Directory.GetFiles(directoryPath, "*.xlsm"))
            {
                logger.Log($"  {Path.GetFileName(file)}");
            }

            var httpClient = new HttpClient();
            var firebaseUrl = "https://enpconventionproject-5ff8f-default-rtdb.firebaseio.com/";
            var firebaseUploader = new FirebaseUploader(httpClient, firebaseUrl, logger);
            var excelProcessor = new ExcuteExcel(firebaseUploader, logger);

            logger.Log($"Processing Excel files in directory: {directoryPath}");
            logger.Log($"Upload type: {uploadType}");

            await excelProcessor.ProcessExcelFiles(directoryPath, uploadType);

            logger.Log("Processing complete.");
        }
        catch (Exception ex)
        {
            logger.Log($"An error occurred: {ex.Message}");
            logger.Log($"Stack Trace: {ex.StackTrace}");
        }

        logger.Log("Press any key to exit...");
        Console.ReadKey();
    }
}
