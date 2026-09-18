using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Readers;

/// <summary>Reads the paragraphs carrying a font style.</summary>
[DisplayName("Read By Font")]
[Description("Reads the paragraphs that are bold, italic or underlined.")]
public sealed class ReadByFont : BaseNativeChild
{
    /// <summary>Which font style to match.</summary>
    [Category("Input")]
    [DisplayName("Font Style")]
    [Description("Which font style to match.")]
    public EnumBoldItalicUnderline FontStyle { get; set; } = EnumBoldItalicUnderline.Bold;

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
        var (array, table) = document.ReadByFont(FontStyle);
        ResultArray.Set(context, array);
        ResultTable.Set(context, table);
    }
}
