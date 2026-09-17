using System.Text;

namespace BalaReva.EasyText.Base;

/// <summary>
/// The text file a <c>TextScope</c> is open on, and the read/write helpers its
/// child activities share.
/// </summary>
/// <remarks>
/// Every read detects the file's encoding and dominant line ending, and every
/// write restores both. Round-tripping a CRLF file through an activity should not
/// silently rewrite it to LF, and a UTF-8-with-BOM file should not lose its BOM —
/// a diff-noise class of bug that is easy to introduce and annoying to trace back
/// to an RPA step.
/// </remarks>
public sealed class TextDocument
{
    /// <summary>Creates a document handle for <paramref name="filePath"/>.</summary>
    public TextDocument(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("The scope's file path is empty.", nameof(filePath));
        FilePath = filePath;
    }

    /// <summary>Full path of the file this scope is open on.</summary>
    public string FilePath { get; }

    /// <summary>Reads the whole file.</summary>
    public string ReadAllText()
    {
        RequireExists();
        return File.ReadAllText(FilePath, DetectEncoding());
    }

    /// <summary>Reads the file as lines, with line endings stripped.</summary>
    public string[] ReadAllLines() => SplitLines(ReadAllText());

    /// <summary>Overwrites the file, preserving its original encoding.</summary>
    public void WriteAllText(string text)
    {
        var encoding = File.Exists(FilePath) ? DetectEncoding() : new UTF8Encoding(false);
        File.WriteAllText(FilePath, text, encoding);
    }

    /// <summary>
    /// Overwrites the file with <paramref name="lines"/>, preserving the original
    /// encoding and line ending, and whether the file ended with a trailing newline.
    /// </summary>
    public void WriteAllLines(IEnumerable<string> lines)
    {
        var original = File.Exists(FilePath) ? File.ReadAllText(FilePath, DetectEncoding()) : string.Empty;
        var newline = DetectNewLine(original);
        var text = string.Join(newline, lines);
        if (EndsWithNewLine(original) && text.Length > 0) text += newline;
        WriteAllText(text);
    }

    /// <summary>Splits text into lines, accepting CRLF, LF and CR.</summary>
    public static string[] SplitLines(string text)
    {
        if (text.Length == 0) return [];
        var lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        // A trailing newline terminates the last line rather than starting a new
        // empty one, so "a\nb\n" is two lines, not three.
        if (lines.Length > 0 && lines[^1].Length == 0) lines = lines[..^1];
        return lines;
    }

    /// <summary>The line ending used by <paramref name="text"/>; CRLF if it has none.</summary>
    public static string DetectNewLine(string text)
    {
        var crlf = 0;
        var lf = 0;
        for (var i = 0; i < text.Length; i++)
        {
            if (text[i] != '\n') continue;
            if (i > 0 && text[i - 1] == '\r') crlf++;
            else lf++;
        }
        if (crlf == 0 && lf == 0) return Environment.NewLine;
        return crlf >= lf ? "\r\n" : "\n";
    }

    private static bool EndsWithNewLine(string text) =>
        text.Length > 0 && (text[^1] == '\n' || text[^1] == '\r');

    private void RequireExists()
    {
        if (!File.Exists(FilePath))
            throw new FileNotFoundException($"The file '{FilePath}' does not exist.", FilePath);
    }

    /// <summary>
    /// Returns the encoding indicated by the file's byte order mark, or UTF-8
    /// without a BOM when there is none.
    /// </summary>
    private Encoding DetectEncoding()
    {
        if (!File.Exists(FilePath)) return new UTF8Encoding(false);

        Span<byte> bom = stackalloc byte[4];
        int read;
        using (var stream = File.OpenRead(FilePath))
        {
            read = stream.Read(bom);
        }

        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return new UTF8Encoding(true);
        if (read >= 4 && bom[0] == 0xFF && bom[1] == 0xFE && bom[2] == 0x00 && bom[3] == 0x00)
            return new UTF32Encoding(false, true);
        if (read >= 4 && bom[0] == 0x00 && bom[1] == 0x00 && bom[2] == 0xFE && bom[3] == 0xFF)
            return new UTF32Encoding(true, true);
        if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE) return new UnicodeEncoding(false, true);
        if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF) return new UnicodeEncoding(true, true);
        return new UTF8Encoding(false);
    }
}
