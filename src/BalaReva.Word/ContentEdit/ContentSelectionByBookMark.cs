using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.ContentEdit;

/// <summary>Moves the selection to a bookmark, page or line.</summary>
[DisplayName("Content Selection By Book Mark")]
[Description("Moves the selection to a bookmark, page or line, optionally extending it.")]
public sealed class ContentSelectionByBookMark : BaseNativeChild
{
    /// <summary>Where to move to.</summary>
    [Category("Input")]
    [DisplayName("Go To")]
    [Description("Whether the destination is a bookmark, a page number or a line number.")]
    public EnumGoTo GoTo { get; set; } = EnumGoTo.Bookmark;

    /// <summary>Bookmark to select.</summary>
    [Category("Input")]
    [DisplayName("Book Mark")]
    [Description("Bookmark to select, when Go To is Bookmark.")]
    public InArgument<string> BookMark { get; set; } = null!;

    /// <summary>Page or line number.</summary>
    [Category("Input")]
    [DisplayName("Go To Page Line No")]
    [Description("Page or line number, when Go To is Page or Line.")]
    public InArgument<int> GoToPageLineNo { get; set; } = null!;

    /// <summary>Lines to extend the selection down by.</summary>
    [Category("Input")]
    [DisplayName("No Rows")]
    [Description("Lines to extend the selection down by. Zero leaves it where it lands.")]
    public InArgument<int> NoRows { get; set; } = null!;

    /// <summary>Characters to extend the selection right by.</summary>
    [Category("Input")]
    [DisplayName("No Columns")]
    [Description("Characters to extend the selection right by. Zero leaves it where it lands.")]
    public InArgument<int> NoColumns { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.Select(
            GoTo,
            BookMark?.Get(context) ?? string.Empty,
            GoToPageLineNo?.Get(context) ?? 0,
            NoRows?.Get(context) ?? 0,
            NoColumns?.Get(context) ?? 0);
}
