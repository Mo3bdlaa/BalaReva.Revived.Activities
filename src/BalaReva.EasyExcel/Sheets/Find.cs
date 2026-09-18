using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Sheets;

/// <summary>Finds text in a range.</summary>
[DisplayName("Find")]
[Description("Finds text in a range.")]
public sealed class Find : BaseActivity
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

    /// <summary>The text of each cell that matched.</summary>
    [Category("Output")]
    [DisplayName("Output")]
    [Description("The text of each cell that matched.")]
    public OutArgument<string[]> Output { get; set; } = null!;

    /// <summary>The rows the matches are on, numbered from 1.</summary>
    [Category("Output")]
    [DisplayName("Row Indexes")]
    [Description("The rows the matches are on, numbered from 1.")]
    public OutArgument<int[]> RowIndexes { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName)
    {
        var (matches, rows) = workbook.Find(
            sheetName,
            CellRange?.Get(context) ?? string.Empty,
            Require(context, FindText, nameof(FindText)),
            FindOption,
            MatchCase);

        Output.Set(context, matches);
        RowIndexes.Set(context, rows);
    }
}
