using System.Text;

public class FirebaseUploader : IFirebaseUploader
{
    private readonly HttpClient _client;
    private readonly string _firebaseUrl;
    private readonly ILogger _logger;

    public FirebaseUploader(HttpClient client, string firebaseUrl, ILogger logger)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _firebaseUrl = firebaseUrl ?? throw new ArgumentNullException(nameof(firebaseUrl));
        _logger = logger ?? throw new ArgumentNullException(nameof(logger));
    }

    public async Task UploadToFirebase(string chapterName, string jsonData)
    {
        try
        {
            var content = new StringContent(jsonData, Encoding.UTF8, "application/json");
            var response = await _client.PutAsync($"{_firebaseUrl}chapters/{chapterName}.json", content);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
                _logger.Log($"Successfully uploaded data for chapter: {chapterName}\nResponse: {responseBody}");
            else
                _logger.Log($"Failed to upload data for chapter: {chapterName}. Status: {response.StatusCode}\nResponse: {responseBody}");
        }
        catch (Exception ex)
        {
            _logger.Log($"Error uploading to Firebase: {ex.Message}\nStack Trace: {ex.StackTrace}");
        }
    }
}