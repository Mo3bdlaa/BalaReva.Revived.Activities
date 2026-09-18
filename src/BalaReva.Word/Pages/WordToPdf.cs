using System.Activities;
using System.ComponentModel;

namespace BalaReva.Word.Pages;

/// <summary>Exports a Word document to PDF.</summary>
/// <remarks>
/// Works on a file directly rather than inside a scope, so it carries its own file and
/// password arguments.
/// </remarks>
[DisplayName("Word To Pdf")]
[Description("Exports a Word document, or a page range of it, to PDF.")]
public sealed class WordToPdf : BaseWord
{
    /// <summary>Path to write the PDF to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("PDF File")]
    [Description("Full path to write the PDF to.")]
    public InArgument<string> PDFFile { get; set; } = null!;

    /// <summary>First page to export.</summary>
    [Category("Input")]
    [DisplayName("Start Page")]
    [Description("First page to export. Zero exports the whole document.")]
    public InArgument<int> StartPage { get; set; } = null!;

    /// <summary>Last page to export.</summary>
    [Category("Input")]
    [DisplayName("End Page")]
    [Description("Last page to export. Zero exports the whole document.")]
    public InArgument<int> EndPage { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordService service) =>
        service.WordToPdf(
            Require(context, WordFile, nameof(WordFile)),
            OpenPassword?.Get(context) ?? string.Empty,
            ModifyPassword?.Get(context) ?? string.Empty,
            Require(context, PDFFile, nameof(PDFFile)),
            StartPage?.Get(context) ?? 0,
            EndPage?.Get(context) ?? 0);
}
