using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;
using BalaReva.EasyText.Utilities;

namespace BalaReva.EasyText.TextFile;

/// <summary>Finds the character offset of a string in the scoped file.</summary>
[DisplayName("Find Text")]
[Description("Finds the character offset of a string within the scoped text file.")]
public sealed class FindText : BaseChildActivity
{
    /// <summary>Text to search for.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find")]
    [Description("Text to search for.")]
    public InArgument<string> Find { get; set; } = null!;

    /// <summary>Character offset to start searching from. 0-based.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Start Index")]
    [Description("Character offset to start searching from. 0-based.")]
    public InArgument<int> StartIndex { get; set; } = null!;

    /// <summary>How to compare text.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text Comparison")]
    [Description("How the search compares text.")]
    public StringComparisonEnum TextComparison { get; set; } = StringComparisonEnum.Ordinal;

    /// <summary>Character offset of the first match, or -1 when not found.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("Character offset of the first match, or -1 when there is no match.")]
    public OutArgument<int> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var find = Find.Get(context);
        if (string.IsNullOrEmpty(find))
            throw new ArgumentException("Find is required.", nameof(Find));

        var text = document.ReadAllText();
        var start = StartIndex.Get(context);
        if (start < 0 || start > text.Length)
        {
            throw new ArgumentOutOfRangeException(
                nameof(StartIndex), start, $"Start Index must be between 0 and {text.Length}.");
        }

        Result.Set(context, text.IndexOf(find, start, TextComparison.ToStringComparison()));
    }
}
