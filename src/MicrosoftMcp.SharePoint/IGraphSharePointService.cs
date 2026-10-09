namespace MicrosoftMcp.SharePoint;

/// <summary>SharePoint sites + explicit-drive file access. Unlike the
/// fixed-drive onedrive_* tools (always the signed-in user's drive), every
/// file operation here takes an explicit driveId: the direct path stays on
/// onedrive_*, the multi-drive path lives here. Teams channel libraries plug
/// in via teams_get_channel_files_folder (driveId + folderId).</summary>
public interface IGraphSharePointService
{
    Task<IReadOnlyList<SharePointSiteInfo>> SearchSitesAsync(
        string query, int top = 25, CancellationToken ct = default);
    Task<SharePointSiteInfo> GetSiteAsync(string siteId, CancellationToken ct = default);
    Task<IReadOnlyList<SharePointDriveInfo>> ListSiteDrivesAsync(
        string siteId, int top = 50, CancellationToken ct = default);
    Task<DriveItemSummary> GetItemAsync(
        string driveId, string itemRef, CancellationToken ct = default);
    Task<IReadOnlyList<DriveItemSummary>> ListChildrenAsync(
        string driveId, string folderRef = "root", int top = 50, CancellationToken ct = default);
    Task<IReadOnlyList<DriveItemSummary>> SearchAsync(
        string driveId, string query, int top = 25, CancellationToken ct = default);
    Task<FileContentDto> DownloadAsync(
        string driveId, string itemRef, int maxBytes = 786432, CancellationToken ct = default);
    /// <summary>Streams a drive file to an absolute host path. Bytes bypass
    /// the model context; the result carries metadata + LocalPath only.</summary>
    Task<FileContentDto> DownloadToFileAsync(
        string driveId, string itemRef, string localPath, bool overwrite = false, CancellationToken ct = default);
    Task<DriveItemSummary> CreateFolderAsync(
        string driveId, string name, string parentRef = "root", CancellationToken ct = default);
    /// <summary>Uploads a file. Either inline content (small text/snippets)
    /// or <paramref name="localPath"/> (absolute host path, read from disk so
    /// bytes bypass the model context). Local files up to 4 MiB use simple
    /// upload, larger ones (up to 100 MiB) a resumable upload session.</summary>
    Task<DriveItemSummary> UploadAsync(
        string driveId,
        string fileName,
        string? parentRef = null,
        string? contentText = null,
        string? contentBase64 = null,
        string? localPath = null,
        CancellationToken ct = default);
    Task<DriveItemSummary> MoveAsync(
        string driveId,
        string itemId,
        string? newParentRef = null,
        string? newName = null,
        CancellationToken ct = default);
}
