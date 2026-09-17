using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Reads the scoped file as an array of lines.</summary>
[DisplayName("Read All Array")]
[Description("Reads the scoped text file as an array of lines.")]
public sealed class ReadAllArray : BaseChildActivity
{
    /// <summary>One entry per line, without line endings.</summary>
    [Category("Output")]
    [DisplayName("Array Result")]
    [Description("One entry per line, with line endings removed.")]
    public OutArgument<string[]> ArrayResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document) =>
        ArrayResult.Set(context, document.ReadAllLines());
}
