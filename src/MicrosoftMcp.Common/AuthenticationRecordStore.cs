using Azure.Identity;
using System.IO.Abstractions;

namespace MicrosoftMcp.Common;

internal sealed class AuthenticationRecordStore
{
    private const string StoreDirectoryName = "microsoft-mcp";
    private const string StoreFileName = "authentication-record.json";
    private readonly IFileSystem _fileSystem;
    private readonly string _path;
    private readonly Action<string, UnixFileMode> _setUnixFileMode;

    internal AuthenticationRecordStore(
        string? path = null,
        IFileSystem? fileSystem = null,
        Action<string, UnixFileMode>? setUnixFileMode = null)
    {
        _path = path ?? GetDefaultPath();
        _fileSystem = fileSystem ?? new FileSystem();
        _setUnixFileMode = setUnixFileMode ?? SetUnixFileMode;
    }

    internal AuthenticationRecord? Load()
    {
        try
        {
            if (!_fileSystem.File.Exists(_path))
            {
                return null;
            }

            using Stream stream = _fileSystem.File.OpenRead(_path);
            return AuthenticationRecord.Deserialize(stream);
        }
        catch (Exception ex) when (ex is IOException
            or UnauthorizedAccessException
            or InvalidDataException
            or FormatException
            or System.Text.Json.JsonException)
        {
            return null;
        }
    }

    internal void Save(AuthenticationRecord record, Action<string> warningSink)
    {
        string? temporaryPath = null;
        try
        {
            string directory = Path.GetDirectoryName(_path)
                ?? throw new InvalidOperationException("The authentication-record path has no directory.");

            bool directoryExisted = _fileSystem.Directory.Exists(directory);
            _fileSystem.Directory.CreateDirectory(directory);
            if (!directoryExisted)
            {
                RestrictToCurrentUser(directory, isDirectory: true);
            }

            temporaryPath = $"{_path}.{Guid.NewGuid():N}.tmp";
            using (Stream stream = _fileSystem.File.Open(temporaryPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                record.Serialize(stream);
                stream.Flush();
            }

            RestrictToCurrentUser(temporaryPath, isDirectory: false);
            _fileSystem.File.Move(temporaryPath, _path, overwrite: true);
            RestrictToCurrentUser(_path, isDirectory: false);
        }
        catch (Exception ex) when (ex is IOException
            or UnauthorizedAccessException
            or InvalidDataException
            or PlatformNotSupportedException)
        {
            warningSink("[auth] Could not persist the authentication record; the next process restart may require interactive login again.");
            if (temporaryPath is not null)
            {
                TryDelete(temporaryPath);
            }
        }
    }

    private static string GetDefaultPath()
    {
        string root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrWhiteSpace(root))
        {
            string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (!string.IsNullOrWhiteSpace(profile))
            {
                root = OperatingSystem.IsWindows()
                    ? Path.Combine(profile, "AppData", "Local")
                    : Path.Combine(profile, ".local", "share");
            }
        }

        if (string.IsNullOrWhiteSpace(root))
        {
            root = AppContext.BaseDirectory;
        }

        return Path.Combine(root, StoreDirectoryName, StoreFileName);
    }

    private void RestrictToCurrentUser(string path, bool isDirectory)
    {
        if (OperatingSystem.IsWindows())
        {
            return;
        }

        UnixFileMode mode = isDirectory
            ? UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute
            : UnixFileMode.UserRead | UnixFileMode.UserWrite;
        _setUnixFileMode(path, mode);
    }

    private static void SetUnixFileMode(string path, UnixFileMode mode)
    {
        if (!OperatingSystem.IsWindows())
        {
            File.SetUnixFileMode(path, mode);
        }
    }

    private void TryDelete(string path)
    {
        try
        {
            _fileSystem.File.Delete(path);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}