using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tables;

/// <summary>Inserts an empty table.</summary>
[DisplayName("Insert Table")]
[Description("Inserts an empty table at a bookmark, page or line.")]
public sealed class InsertTable : BaseNativeChild
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

    /// <summary>How many rows.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("No Rows")]
    [Description("How many rows the new table has.")]
    public InArgument<int> NoRows { get; set; } = null!;

    /// <summary>How many columns.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("No Columns")]
    [Description("How many columns the new table has.")]
    public InArgument<int> NoColumns { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.InsertTable(
            GoTo,
            BookMark?.Get(context) ?? string.Empty,
            GoToPageLineNo?.Get(context) ?? 0,
            NoRows.Get(context),
            NoColumns.Get(context));
}
