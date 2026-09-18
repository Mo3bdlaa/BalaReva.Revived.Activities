using System.Activities;
using System.ComponentModel;
using System.Data;

namespace BalaReva.Word.Readers;

/// <summary>Reads the paragraphs carrying a paragraph style.</summary>
[DisplayName("Read By Style")]
[Description("Reads the paragraphs that carry a named paragraph style.")]
public sealed class ReadByStyle : BaseNativeChild
{
    /// <summary>Paragraph style to match.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Paragraph Style")]
    [Description("Name of the paragraph style to match, for example Heading 1.")]
    public InArgument<string> ParagraphStyle { get; set; } = null!;

    /// <summary>The matching paragraphs.</summary>
    [Category("Output")]
    [DisplayName("Result Array")]
    [Description("The matching paragraphs, one entry each.")]
    public OutArgument<string[]> ResultArray { get; set; } = null!;

    /// <summary>The matching paragraphs, as a single-column table.</summary>
    [Category("Output")]
    [DisplayName("Result Table")]
    [Description("The matching paragraphs, as a single-column DataTable.")]
    public OutArgument<DataTable> ResultTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document)
    {
        var (array, table) = document.ReadByStyle(Require(context, ParagraphStyle, nameof(ParagraphStyle)));
        ResultArray.Set(context, array);
        ResultTable.Set(context, table);
    }
}
