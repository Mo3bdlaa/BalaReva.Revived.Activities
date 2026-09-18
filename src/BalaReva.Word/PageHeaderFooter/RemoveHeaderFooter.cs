using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.PageHeaderFooter;

/// <summary>Clears the header or footer of every section.</summary>
[DisplayName("Remove Header Footer")]
[Description("Clears the header or footer of every section in the document.")]
public sealed class RemoveHeaderFooter : BaseNativeChild
{
    /// <summary>Whether to clear headers or footers.</summary>
    [Category("Input")]
    [DisplayName("Remove Type")]
    [Description("Whether to clear headers or footers.")]
    public EnumHeadersFooters RemoveType { get; set; } = EnumHeadersFooters.Header;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.RemoveHeaderFooter(RemoveType);
}
