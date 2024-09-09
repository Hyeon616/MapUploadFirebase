public interface IFirebaseUploader
{
    Task UploadToFirebaseChapter(string chapterName, string jsonData);
    Task UploadToFirebaseAnswers(string chapterName, string jsonData);
}
