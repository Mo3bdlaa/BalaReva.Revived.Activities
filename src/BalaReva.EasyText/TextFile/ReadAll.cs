using System.Activities;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.TextFile;

/// <summary>Reads the whole scoped file into a single string.</summary>
[DisplayName("Read All")]
[Description("Reads the entire contents of the scoped text file.")]
public sealed class ReadAll : BaseChildActivity
{
    /// <summary>The file's full contents.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The entire contents of the file.")]
    public OutArgument<string> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteChild(CodeActivityContext context, TextDocument document) =>
        Result.Set(context, document.ReadAllText());
}
