public interface IFirebaseUploader
{
    Task UploadToFirebase(string chapterName, string jsonData);
}
