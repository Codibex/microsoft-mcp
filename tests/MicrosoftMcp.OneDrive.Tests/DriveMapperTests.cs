using AwesomeAssertions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.OneDrive.Tests;

public sealed class DriveMapperTests
{
    [Fact]
    public void MapItem_is_null_safe()
    {
        var summary = DriveMapper.MapItem(new DriveItem());

        summary.Id.Should().BeEmpty();
        summary.IsFolder.Should().BeFalse();
        summary.ChildCount.Should().Be(0);
    }

    [Fact]
    public void MapItem_detects_folders_and_files()
    {
        var folder = DriveMapper.MapItem(new DriveItem
        {
            Id = "f1",
            Name = "Docs",
            Folder = new Folder { ChildCount = 3 }
        });
        folder.IsFolder.Should().BeTrue();
        folder.ChildCount.Should().Be(3);

        var file = DriveMapper.MapItem(new DriveItem
        {
            Id = "f2",
            Name = "a.txt",
            File = new FileObject { MimeType = "text/plain" }
        });
        file.IsFolder.Should().BeFalse();
        file.MimeType.Should().Be("text/plain");
    }

    [Fact]
    public void MapDrive_maps_quota()
    {
        var drive = DriveMapper.MapDrive(new Drive
        {
            Id = "d1",
            Name = "OneDrive",
            Quota = new Quota { Total = 100, Used = 40, Remaining = 60 }
        });

        drive.TotalBytes.Should().Be(100);
        drive.RemainingBytes.Should().Be(60);
    }
}
