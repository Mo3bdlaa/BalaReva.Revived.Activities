using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Sets the workbook's open and modify passwords.</summary>
[DisplayName("Set Password")]
[Description("Sets the workbook's open and modify passwords.")]
public sealed class SetPassword : ExcelCore
{
    /// <summary>New password for opening the workbook.</summary>
    [Category("Input")]
    [DisplayName("New Password")]
    [Description("New password for opening the workbook. Empty removes it.")]
    public InArgument<string> NewPassword { get; set; } = null!;

    /// <summary>New password for changing the workbook.</summary>
    [Category("Input")]
    [DisplayName("New Modify Password")]
    [Description("New password for changing the workbook. Empty removes it.")]
    public InArgument<string> NewModifyPassword { get; set; } = null!;

    /// <summary>True when the passwords were set.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the passwords were set.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.SetPassword(
            NewPassword?.Get(context) ?? string.Empty,
            NewModifyPassword?.Get(context) ?? string.Empty);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
