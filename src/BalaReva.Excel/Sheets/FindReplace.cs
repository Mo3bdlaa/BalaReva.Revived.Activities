using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;
using BalaReva.Excel.Enums;

namespace BalaReva.Excel.Sheets;

/// <summary>Finds and replaces text on a sheet.</summary>
[DisplayName("Find Replace")]
[Description("Finds text on a sheet and replaces it.")]
public sealed class FindReplace : ExcelCore
{
    /// <summary>Text to find.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Find")]
    [Description("Text to find.")]
    public InArgument<string> Find { get; set; } = null!;

    /// <summary>Text to put in its place.</summary>
    [Category("Input")]
    [DisplayName("Replace")]
    [Description("Text to put in its place. Empty clears the match.")]
    public InArgument<string> Replace { get; set; } = null!;

    /// <summary>Range to search. Empty searches the whole used range.</summary>
    [Category("Input")]
    [DisplayName("Cell Range")]
    [Description("Range to search. Empty searches the whole used range.")]
    public InArgument<string> CellRange { get; set; } = null!;

    /// <summary>Whether to match whole cells only.</summary>
    [Category("Input")]
    [DisplayName("Find Option")]
    [Description("Whole matches the entire cell; Part matches within it.")]
    public FindReplaceEnum FindOption { get; set; } = FindReplaceEnum.Part;

    /// <summary>True when the replace ran.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the replace ran.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.FindReplace(
            sheetName,
            CellRange?.Get(context) ?? string.Empty,
            Require(context, Find, nameof(Find)),
            Replace?.Get(context) ?? string.Empty,
            FindOption);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
