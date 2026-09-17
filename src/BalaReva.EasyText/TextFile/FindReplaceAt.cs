using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Replaces a string, but only within a range of lines.</summary>
[DisplayName("Find Replace At")]
[Description("Replaces every occurrence of a string within an inclusive range of lines.")]
public sealed class FindReplaceAt : BaseChildActivity
{
    /// <summary>First line to replace within, 1-based and inclusive.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Line From")]
    [Description("First line to replace within. 1-based and inclusive.")]
    public InArgument<int> LineFrom { get; set; } = null!;

    /// <summary>Last line to replace within, 1-based and inclusive.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Line To")]
    [Description("Last line to replace within. 1-based and inclusive.")]
    public InArgument<int> LineTo { get; set; } = null!;

    /// <summary>Text to replace.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Old Text")]
    [Description("Text to replace.")]
    public InArgument<string> OldText { get; set; } = null!;

    /// <summary>Replacement text.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New Text")]
    [Description("Text to put in its place.")]
    public InArgument<string> NewText { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var oldText = OldText.Get(context);
        if (string.IsNullOrEmpty(oldText))
            throw new ArgumentException("Old Text is required.", nameof(OldText));
        var newText = NewText.Get(context) ?? string.Empty;

        var lines = document.ReadAllLines();
        var from = LineIndex(LineFrom.Get(context), lines.Length, nameof(LineFrom));
        var to = LineIndex(LineTo.Get(context), lines.Length, nameof(LineTo));
        if (from > to)
            throw new ArgumentException("Line From must not be greater than Line To.", nameof(LineFrom));

        for (var i = from; i <= to; i++)
            lines[i] = lines[i].Replace(oldText, newText, StringComparison.Ordinal);

        document.WriteAllLines(lines);
    }
}
