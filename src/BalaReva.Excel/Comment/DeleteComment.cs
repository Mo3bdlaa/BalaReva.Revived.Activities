using System.Activities;
using System.ComponentModel;

namespace BalaReva.Excel.Comment;

/// <summary>Removes a cell's comment.</summary>
[DisplayName("Delete Comment")]
[Description("Removes a cell's comment.")]
public sealed class DeleteComment : BaseCommnet
{
    /// <summary>True when the comment was removed.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the comment was removed.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.DeleteComment(sheetName, Require(context, Cell, nameof(Cell)));

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
