using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Exports the presentation to PDF.</summary>
[DisplayName("Export Pdf")]
[Description("Exports the presentation to PDF.")]
public sealed class ExportPdf : BaseNativeChild
{
    /// <summary>Full path to write the PDF to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path to write the PDF to.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Whether the PDF is meant for screen or for print.</summary>
    [Category("Input")]
    [DisplayName("Format Type")]
    [Description("Whether the PDF is meant for screen or for print.")]
    public FixedFormatIntentEnum FormatType { get; set; } = FixedFormatIntentEnum.Print;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.ExportPdf(Require(context, FilePath, nameof(FilePath)), FormatType);
}
