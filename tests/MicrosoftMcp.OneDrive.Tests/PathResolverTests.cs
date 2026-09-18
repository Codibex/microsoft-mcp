using AwesomeAssertions;

namespace MicrosoftMcp.OneDrive.Tests;

public sealed class PathResolverTests
{
    [Theory]
    [InlineData("/a/b", true)]
    [InlineData("  /a", true)]
    [InlineData("root", false)]
    [InlineData("abc123", false)]
    public void Detects_paths(string input, bool expected) =>
        PathResolver.IsPath(input).Should().Be(expected);

    [Theory]
    [InlineData("root", true)]
    [InlineData("ROOT", true)]
    [InlineData("abc", false)]
    public void Detects_root(string input, bool expected) =>
        PathResolver.IsRoot(input).Should().Be(expected);

    [Fact]
    public void Normalizes_path() =>
        PathResolver.NormalizePath("  /a/b ").Should().Be("a/b");

    [Fact]
    public void Blank_ref_throws_coded_error() =>
        Assert.Throws<MicrosoftMcp.Common.GraphServiceException>(
            () => PathResolver.RequireRef("  ")).Code.Should().Be("invalid-request");
}
