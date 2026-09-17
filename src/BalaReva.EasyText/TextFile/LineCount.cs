using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Counts the lines in the scoped file.</summary>
[DisplayName("Line Count")]
[Description("Counts the lines in the scoped text file.")]
public sealed class LineCount : BaseChildActivity
{
    /// <summary>Number of lines in the file.</summary>
    [RequiredArgument]
    [Category("Output")]
    [DisplayName("Result")]
    [Description("Number of lines in the file.")]
    public OutArgument<int> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document) =>
        Result.Set(context, document.ReadAllLines().Length);
}
