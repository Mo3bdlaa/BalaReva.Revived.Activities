using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Inserts a table filled from a DataTable.</summary>
[DisplayName("Insert Data Table")]
[Description("Inserts a table filled from a DataTable, at a bookmark, page or line.")]
public sealed class InsertDataTable : BaseNativeChild
{
    /// <summary>Where to insert.</summary>
    [Category("Input")]
    [DisplayName("Go To")]
    [Description("Whether the destination is a bookmark, a page number or a line number.")]
    public EnumGoTo GoTo { get; set; } = EnumGoTo.Bookmark;

    /// <summary>Bookmark to insert at.</summary>
    [Category("Input")]
    [DisplayName("Book Mark")]
    [Description("Bookmark to insert at, when Go To is Bookmark.")]
    public InArgument<string> BookMark { get; set; } = null!;

    /// <summary>Page or line number to insert at.</summary>
    [Category("Input")]
    [DisplayName("Go To Page Line No")]
    [Description("Page or line number to insert at, when Go To is Page or Line.")]
    public InArgument<int> GoToPageLineNo { get; set; } = null!;

    /// <summary>Data to write into the table.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Input Table")]
    [Description("Data to write into the new table.")]
    public InArgument<DataTable> InputTable { get; set; } = null!;

    /// <summary>Whether to write the column names as a first row.</summary>
    [Category("Input")]
    [DisplayName("Add Header")]
    [Description("Write the column names as the table's first row.")]
    public bool AddHeader { get; set; }

    /// <summary>Word table style to apply.</summary>
    [Category("Input")]
    [DisplayName("Style Name")]
    [Description("Name of a Word table style to apply. Empty leaves the default.")]
    public InArgument<string> StyleName { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document)
    {
        var input = InputTable?.Get(context)
            ?? throw new ArgumentException("InputTable is required.", nameof(InputTable));

        document.InsertDataTable(
            GoTo,
            BookMark?.Get(context) ?? string.Empty,
            GoToPageLineNo?.Get(context) ?? 0,
            input,
            AddHeader,
            StyleName?.Get(context) ?? string.Empty);
    }
}
