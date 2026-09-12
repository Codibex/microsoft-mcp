using AwesomeAssertions;
using Microsoft.Graph.Models;

namespace MicrosoftMcp.Outlook.Tests;

public sealed class EmailMapperTests
{
    [Fact]
    public void Truncate_keeps_short_strings_and_marks_long_ones()
    {
        EmailMapper.Truncate(null, 5).Should().BeNull();
        EmailMapper.Truncate("abc", 5).Should().Be("abc");
        EmailMapper.Truncate("abcdef", 5).Should().Be("abcde…[truncated]");
    }

    [Fact]
    public void MapSummary_is_null_safe_and_truncates_preview()
    {
        var summary = EmailMapper.MapSummary(new Message { BodyPreview = new string('x', 400) });

        summary.Id.Should().BeEmpty();
        summary.Subject.Should().BeEmpty();
        summary.From.Should().BeNull();
        summary.IsRead.Should().BeFalse();
        summary.Preview.Should().HaveLength(300 + "…[truncated]".Length);
    }

    [Fact]
    public void MapDetail_maps_addresses_and_truncates_body()
    {
        var detail = EmailMapper.MapDetail(new Message
        {
            Id = "m1",
            Subject = "Hi",
            From = new Recipient { EmailAddress = new EmailAddress { Name = "A", Address = "a@x.y" } },
            ToRecipients = [new Recipient { EmailAddress = new EmailAddress { Address = "b@x.y" } }],
            Body = new ItemBody { Content = new string('y', 9000) }
        });

        detail.From!.Address.Should().Be("a@x.y");
        detail.ToRecipients.Should().HaveCount(1);
        detail.Body.Should().HaveLength(8000 + "…[truncated]".Length);
    }

    [Fact]
    public void MergeCategories_adds_removes_and_dedups_case_insensitive()
    {
        var merged = EmailMapper.MergeCategories(["Red", "Blue"], ["blue", "Green"], ["RED"]);

        merged.Should().BeEquivalentTo(["Blue", "Green"]);
    }

    [Fact]
    public void MapAttachment_maps_fields()
    {
        var info = EmailMapper.MapAttachment(new FileAttachment
        {
            Id = "a1",
            Name = "doc.pdf",
            ContentType = "application/pdf",
            Size = 42,
            IsInline = false
        });

        info.Should().Be(new AttachmentInfo("a1", "doc.pdf", "application/pdf", 42, false));
    }
}
