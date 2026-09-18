using System.Activities;
using System.ComponentModel;

namespace BalaReva.Excel.Comment;

/// <summary>Reads a cell's comment.</summary>
/// <remarks>
/// Reports the comment text rather than an ExecutionResult, which is the one comment
/// activity shaped that way in the published package.
/// </remarks>
[DisplayName("Get Comment")]
[Description("Reads a cell's comment, or empty when it has none.")]
public sealed class GetComment : BaseCommnet
{
    /// <summary>The comment text.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The comment text, or empty when the cell has none.")]
    public OutArgument<string> Result { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        Result.Set(context, workbook.GetComment(sheetName, Require(context, Cell, nameof(Cell))));
}
