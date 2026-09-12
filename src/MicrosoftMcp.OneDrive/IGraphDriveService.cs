namespace MicrosoftMcp.OneDrive;

public interface IGraphDriveService
{
    Task<DriveInfoDto> GetDriveAsync(CancellationToken ct = default);
    Task<IReadOnlyList<DriveInfoDto>> ListDrivesAsync(CancellationToken ct = default);
    Task<DriveItemSummary> GetItemAsync(string itemRef, CancellationToken ct = default);
    Task<IReadOnlyList<DriveItemSummary>> ListChildrenAsync(
        string folderRef = "root", int top = 50, CancellationToken ct = default);
    Task<IReadOnlyList<DriveItemSummary>> SearchAsync(
        string query, int top = 25, CancellationToken ct = default);
    Task<FileContentDto> DownloadAsync(
        string itemRef, int maxBytes = 786432, CancellationToken ct = default);
    Task<DriveItemSummary> CreateFolderAsync(
        string name, string parentRef = "root", CancellationToken ct = default);
    Task<DriveItemSummary> UploadAsync(
        string fileName,
        string? parentRef = null,
        string? contentText = null,
        string? contentBase64 = null,
        CancellationToken ct = default);
    Task<DriveItemSummary> MoveAsync(
        string itemId,
        string? newParentRef = null,
        string? newName = null,
        CancellationToken ct = default);
}
