using System.Activities;
using System.ComponentModel;

namespace BalaReva.Excel.Comment;

/// <summary>Attaches a comment to a cell.</summary>
[DisplayName("Add Comment")]
[Description("Attaches a comment to a cell, replacing any comment already there.")]
public sealed class AddComment : BaseCommnet
{
    /// <summary>Text of the comment.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Comment")]
    [Description("Text of the comment.")]
    public InArgument<string> Comment { get; set; } = null!;

    /// <summary>True when the comment was written.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the comment was written.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.AddComment(
            sheetName,
            Require(context, Cell, nameof(Cell)),
            Comment?.Get(context) ?? string.Empty);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
