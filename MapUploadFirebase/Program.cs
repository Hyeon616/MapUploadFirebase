class Program
{
    public static async Task Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the directory path as an argument.");
            return;
        }

        var logger = new ConsoleLogger();
        var httpClient = new HttpClient();
        var firebaseUrl = "https://enpconventionproject-5ff8f-default-rtdb.firebaseio.com/";
        var firebaseUploader = new FirebaseUploader(httpClient, firebaseUrl, logger);
        var excelProcessor = new ExcuteExcel(firebaseUploader, logger);

        try
        {
            await excelProcessor.ProcessExcelFiles(args[0]);
        }
        catch (Exception ex)
        {
            logger.Log($"An error occurred: {ex.Message}");
            logger.Log($"Stack Trace: {ex.StackTrace}");
        }

        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
