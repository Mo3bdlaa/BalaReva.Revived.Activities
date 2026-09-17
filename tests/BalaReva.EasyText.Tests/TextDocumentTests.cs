using System.Text;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.Tests;

public class TextDocumentTests
{
    [Theory]
    [InlineData("a\nb\nc", new[] { "a", "b", "c" })]
    [InlineData("a\r\nb\r\nc", new[] { "a", "b", "c" })]
    [InlineData("a\rb\rc", new[] { "a", "b", "c" })]
    // A trailing newline terminates the last line; it does not start an empty one.
    [InlineData("a\nb\n", new[] { "a", "b" })]
    [InlineData("", new string[0])]
    [InlineData("\n", new[] { "" })]
    public void SplitLines_handles_every_line_ending(string text, string[] expected) =>
        Assert.Equal(expected, TextDocument.SplitLines(text));

    [Theory]
    [InlineData("a\r\nb", "\r\n")]
    [InlineData("a\nb", "\n")]
    // Mixed endings resolve to whichever dominates.
    [InlineData("a\r\nb\r\nc\nd", "\r\n")]
    [InlineData("a\nb\nc\r\nd", "\n")]
    public void DetectNewLine_picks_the_dominant_ending(string text, string expected) =>
        Assert.Equal(expected, TextDocument.DetectNewLine(text));

    [Fact]
    public void Writing_preserves_CRLF_endings()
    {
        using var file = new ScopedFile("one\r\ntwo\r\nthree");
        var document = new TextDocument(file.Path);

        document.WriteAllLines(["one", "two"]);

        Assert.Equal("one\r\ntwo", file.Contents);
    }

    [Fact]
    public void Writing_preserves_LF_endings()
    {
        using var file = new ScopedFile("one\ntwo\nthree");
        var document = new TextDocument(file.Path);

        document.WriteAllLines(["one", "two"]);

        Assert.Equal("one\ntwo", file.Contents);
    }

    [Fact]
    public void Writing_preserves_a_trailing_newline()
    {
        using var file = new ScopedFile("one\ntwo\n");
        var document = new TextDocument(file.Path);

        document.WriteAllLines(["one", "two", "three"]);

        Assert.Equal("one\ntwo\nthree\n", file.Contents);
    }

    [Fact]
    public void Writing_preserves_a_UTF8_byte_order_mark()
    {
        var path = Path.Combine(Path.GetTempPath(), $"easytext-bom-{Guid.NewGuid():N}.txt");
        File.WriteAllText(path, "one\ntwo", new UTF8Encoding(true));
        try
        {
            new TextDocument(path).WriteAllLines(["one", "two", "three"]);

            var bytes = File.ReadAllBytes(path);
            Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, bytes[..3]);
            Assert.Equal("one\ntwo\nthree", File.ReadAllText(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Reading_a_file_that_is_not_there_names_the_file()
    {
        var missing = Path.Combine(Path.GetTempPath(), $"absent-{Guid.NewGuid():N}.txt");
        var document = new TextDocument(missing);

        var error = Assert.Throws<FileNotFoundException>(() => document.ReadAllText());
        Assert.Contains(missing, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void An_empty_path_is_rejected_at_construction() =>
        Assert.Throws<ArgumentException>(() => new TextDocument("  "));
}
