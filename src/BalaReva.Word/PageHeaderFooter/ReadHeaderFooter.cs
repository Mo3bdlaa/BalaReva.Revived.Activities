using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.PageHeaderFooter;

/// <summary>Reads the header or footer text of every section.</summary>
[DisplayName("Read Header Footer")]
[Description("Reads the header or footer text of every section in the document.")]
public sealed class ReadHeaderFooter : BaseNativeChild
{
    /// <summary>Whether to read headers or footers.</summary>
    [Category("Input")]
    [DisplayName("Read Type")]
    [Description("Whether to read headers or footers.")]
    public EnumHeadersFooters ReadType { get; set; } = EnumHeadersFooters.Header;

    /// <summary>The text found, one entry per non-empty header or footer.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The text found, one entry per non-empty header or footer.")]
    public OutArgument<string[]> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        Result.Set(context, document.ReadHeaderFooter(ReadType));
}
