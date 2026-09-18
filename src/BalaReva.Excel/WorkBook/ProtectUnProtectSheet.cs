using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.WorkBook;

/// <summary>Protects or unprotects a sheet.</summary>
[DisplayName("Protect Un Protect Sheet")]
[Description("Protects or unprotects a sheet with a password.")]
public sealed class ProtectUnProtectSheet : ExcelCore
{
    /// <summary>Password to protect with, or to unprotect using.</summary>
    [Category("Input")]
    [DisplayName("Protect Password")]
    [Description("Password to protect the sheet with, or to unprotect it using.")]
    public InArgument<string> ProtectPassword { get; set; } = null!;

    /// <summary>Whether to protect or unprotect.</summary>
    [Category("Input")]
    [DisplayName("Protect Type")]
    [Description("Whether to protect or unprotect the sheet.")]
    public ProtectUnProtectEnum ProtectType { get; set; } = ProtectUnProtectEnum.Protect;

    /// <summary>True when the sheet's protection was set.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the sheet's protection was set.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.ProtectSheet(sheetName, ProtectPassword?.Get(context) ?? string.Empty, ProtectType);

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
