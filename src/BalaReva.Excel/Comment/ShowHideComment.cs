using System.Activities;
using System.ComponentModel;

namespace BalaReva.Excel.Comment;

/// <summary>Shows or hides a cell's comment.</summary>
[DisplayName("Show Hide Comment")]
[Description("Shows or hides a cell's comment.")]
public sealed class ShowHideComment : BaseCommnet
{
    /// <summary>Whether to show the comment.</summary>
    [Category("Input")]
    [DisplayName("Show Comment")]
    [Description("True shows the comment, false hides it.")]
    public bool ShowComment { get; set; }

    /// <summary>True when the comment's visibility was set.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the comment's visibility was set.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.ShowHideComment(sheetName, Require(context, Cell, nameof(Cell)), ShowComment);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
