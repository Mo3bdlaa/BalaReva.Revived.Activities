using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;
using BalaReva.EasyText.Utilities;

namespace BalaReva.EasyText.TextFile;

/// <summary>Finds every line of the scoped file that contains a string.</summary>
[DisplayName("Find Line Index")]
[Description("Finds the line numbers of every line containing the given text.")]
public sealed class FindLineIndex : BaseChildActivity
{
    /// <summary>Text to search for.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find")]
    [Description("Text to search for.")]
    public InArgument<string> Find { get; set; } = null!;

    /// <summary>How to compare text.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Text Comparison")]
    [Description("How the search compares text.")]
    public StringComparisonEnum TextComparison { get; set; } = StringComparisonEnum.Ordinal;

    /// <summary>1-based line numbers of every matching line; empty when none match.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("1-based line numbers of every matching line. Empty when nothing matches.")]
    public OutArgument<int[]> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document)
    {
        var find = Find.Get(context);
        if (string.IsNullOrEmpty(find))
            throw new ArgumentException("Find is required.", nameof(Find));

        var comparison = TextComparison.ToStringComparison();
        var lines = document.ReadAllLines();
        var matches = new List<int>();
        for (var i = 0; i < lines.Length; i++)
        {
            if (lines[i].Contains(find, comparison)) matches.Add(i + 1);
        }

        Result.Set(context, matches.ToArray());
    }
}
