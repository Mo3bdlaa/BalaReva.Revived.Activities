using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.PageHeaderFooter;

/// <summary>Saves header or footer images to a folder.</summary>
[DisplayName("Extract Header Footer Images")]
[Description("Saves the images in the document's headers or footers into a folder.")]
public sealed class ExtractHeaderFooterImages : BaseNativeChild
{
    /// <summary>Whether to read headers or footers.</summary>
    [Category("Input")]
    [DisplayName("Read Type")]
    [Description("Whether to read headers or footers.")]
    public EnumHeadersFooters ReadType { get; set; } = EnumHeadersFooters.Header;

    /// <summary>Folder to save the images into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Save Folder")]
    [Description("Folder to save the images into. Created if it does not exist.")]
    public InArgument<string> SaveFolder { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.ExtractHeaderFooterImages(ReadType, Require(context, SaveFolder, nameof(SaveFolder)));
}
