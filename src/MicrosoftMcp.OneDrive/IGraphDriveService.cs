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
    /// <summary>Streams a drive file to an absolute host path. Bytes bypass
    /// the model context; the result carries metadata + LocalPath only.</summary>
    Task<FileContentDto> DownloadToFileAsync(
        string itemRef, string localPath, bool overwrite = false, CancellationToken ct = default);
    Task<DriveItemSummary> CreateFolderAsync(
        string name, string parentRef = "root", CancellationToken ct = default);
    /// <summary>Uploads a file. Either inline content (small text/snippets)
    /// or <paramref name="localPath"/> (absolute host path, read from disk so
    /// bytes bypass the model context). Local files up to 4 MiB use simple
    /// upload, larger ones (up to 100 MiB) a resumable upload session.</summary>
    Task<DriveItemSummary> UploadAsync(
        string fileName,
        string? parentRef = null,
        string? contentText = null,
        string? contentBase64 = null,
        string? localPath = null,
        CancellationToken ct = default);
    Task<DriveItemSummary> MoveAsync(
        string itemId,
        string? newParentRef = null,
        string? newName = null,
        CancellationToken ct = default);
}
