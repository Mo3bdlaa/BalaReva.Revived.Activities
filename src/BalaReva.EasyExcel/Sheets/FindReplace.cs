using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Replaces text in a range.</summary>
[DisplayName("Find Replace")]
[Description("Replaces text in a range.")]
public sealed class FindReplace : BaseActivity
{
    /// <summary>Range to search. Empty means the whole sheet.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to search. Empty means the whole sheet.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Match the whole cell or any part of it.</summary>
    [Category("Input")]
    [DisplayName("Find Option")]
    [Description("Match the whole cell or any part of it.")]
    public FindReplaceEnum FindOption { get; set; } = FindReplaceEnum.Whole;

    /// <summary>Text to look for.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find Text")]
    [Description("Text to look for.")]
    public InArgument<string> FindText { get; set; } = null!;

    /// <summary>Whether the search is case sensitive.</summary>
    [Category("Input")]
    [DisplayName("Match Case")]
    [Description("Whether the search is case sensitive.")]
    public bool MatchCase { get; set; }

    /// <summary>True when something was replaced.</summary>
    [Category("Output")]
    [DisplayName("Output")]
    [Description("True when something was replaced.")]
    public OutArgument<bool> Output { get; set; } = null!;

    /// <summary>Text to put in its place.</summary>
    [Category("Input")]
    [DisplayName("Replace Text")]
    [Description("Text to put in its place.")]
    public InArgument<string> ReplaceText { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
        => Output.Set(context, workbook.FindReplace(
            sheetName,
            CellRange?.Get(context) ?? string.Empty,
            Require(context, FindText, nameof(FindText)),
            ReplaceText?.Get(context) ?? string.Empty,
            FindOption,
            MatchCase));
}
