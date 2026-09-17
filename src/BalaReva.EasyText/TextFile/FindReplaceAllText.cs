using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Replaces every occurrence of a string throughout the scoped file.</summary>
[DisplayName("Find Replace All Text")]
[Description("Replaces every occurrence of a string throughout the scoped text file.")]
public sealed class FindReplaceAllText : BaseChildActivity
{
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

        var text = document.ReadAllText();
        document.WriteAllText(text.Replace(oldText, NewText.Get(context) ?? string.Empty, StringComparison.Ordinal));
    }
}
