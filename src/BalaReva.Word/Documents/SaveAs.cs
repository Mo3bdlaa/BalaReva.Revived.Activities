using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Documents;

/// <summary>Saves a copy of the document under another name and format.</summary>
[DisplayName("Save As")]
[Description("Saves a copy of the document under another name and format.")]
public sealed class SaveAs : BaseNativeChild
{
    /// <summary>Path to save to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New File Name")]
    [Description("Full path to save the copy to.")]
    public InArgument<string> NewFileName { get; set; } = null!;

    /// <summary>Format to save in.</summary>
    [Category("Input")]
    [DisplayName("Save As Format")]
    [Description("Format to save the copy in.")]
    public EnumSaveAs SaveAsFormat { get; set; } = EnumSaveAs.Document;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.SaveAs(Require(context, NewFileName, nameof(NewFileName)), SaveAsFormat);
}
