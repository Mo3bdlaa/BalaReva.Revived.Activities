using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Reads an inclusive range of lines from the scoped file.</summary>
[DisplayName("Read Specific Line")]
[Description("Reads an inclusive range of lines from the scoped text file.")]
public sealed class ReadSpecificLine : BaseChildActivity
{
    /// <summary>First line to read, 1-based and inclusive.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Line From")]
    [Description("First line to read. 1-based and inclusive.")]
    public InArgument<int> LineFrom { get; set; } = null!;

    /// <summary>Last line to read, 1-based and inclusive.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Line To")]
    [Description("Last line to read. 1-based and inclusive.")]
    public InArgument<int> LineTo { get; set; } = null!;

    /// <summary>The selected lines joined by the file's line ending.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The selected lines, joined by the file's line ending.")]
    public OutArgument<string> Result { get; set; } = null!;

    /// <summary>The selected lines, one array entry each.</summary>
    [Category("Output")]
    [DisplayName("Array Result")]
    [Description("The selected lines, one array entry each.")]
    public OutArgument<string[]> ArrayResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var text = document.ReadAllText();
        var lines = TextDocument.SplitLines(text);

        var from = LineIndex(LineFrom.Get(context), lines.Length, nameof(LineFrom));
        var to = LineIndex(LineTo.Get(context), lines.Length, nameof(LineTo));
        if (from > to)
            throw new ArgumentException("Line From must not be greater than Line To.", nameof(LineFrom));

        var selected = lines[from..(to + 1)];
        ArrayResult.Set(context, selected);
        Result.Set(context, string.Join(TextDocument.DetectNewLine(text), selected));
    }
}
