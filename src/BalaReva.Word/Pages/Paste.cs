using System.Activities;
using System.ComponentModel;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Pages;

/// <summary>Pastes the clipboard into the document.</summary>
[DisplayName("Paste")]
[Description("Pastes the clipboard contents at a bookmark, page or line.")]
public sealed class Paste : BaseNativeChild
{
    /// <summary>Where to paste.</summary>
    [Category("Input")]
    [DisplayName("Go To")]
    [Description("Whether the destination is a bookmark, a page number or a line number.")]
    public EnumGoTo GoTo { get; set; } = EnumGoTo.Bookmark;

    /// <summary>Bookmark to paste at, when Go To is Bookmark.</summary>
    [Category("Input")]
    [DisplayName("Book Mark")]
    [Description("Bookmark to paste at, when Go To is Bookmark.")]
    public InArgument<string> BookMark { get; set; } = null!;

    /// <summary>Page or line number, when Go To is Page or Line.</summary>
    [Category("Input")]
    [DisplayName("Go To Page Line No")]
    [Description("Page or line number to paste at, when Go To is Page or Line.")]
    public InArgument<int> GoToPageLineNo { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IWordDocument document) =>
        document.Paste(GoTo, BookMark?.Get(context) ?? string.Empty, GoToPageLineNo?.Get(context) ?? 0);
}
