using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel.Sheets;

/// <summary>Copies a sheet into another workbook.</summary>
[DisplayName("Copy To File")]
[Description("Copies a sheet into another workbook.")]
public sealed class CopyToFile : ExcelCore
{
    /// <summary>Workbook to copy the sheet into.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("New File Path")]
    [Description("Full path of the workbook to copy the sheet into.")]
    public InArgument<string> NewFilePath { get; set; } = null!;

    /// <summary>Name to give the copied sheet.</summary>
    [Category("Input")]
    [DisplayName("New Sheet Name")]
    [Description("Name to give the copied sheet. Empty lets Excel name it.")]
    public InArgument<string> NewSheetName { get; set; } = null!;

    /// <summary>Password needed to open the destination.</summary>
    [Category("Input")]
    [DisplayName("New File Password")]
    [Description("Password needed to open the destination workbook, if it has one.")]
    public InArgument<string> NewFilePassword { get; set; } = null!;

    /// <summary>Password needed to change the destination.</summary>
    [Category("Input")]
    [DisplayName("New Modify File Password")]
    [Description("Password needed to change the destination workbook, if it has one.")]
    public InArgument<string> NewModifyFilePassword { get; set; } = null!;

    /// <summary>Whether to create the destination when it is missing.</summary>
    [Category("Input")]
    [DisplayName("Auto File Creation")]
    [Description("Create the destination workbook when it does not exist.")]
    public InArgument<bool> AutoFileCreation { get; set; } = null!;

    /// <summary>True when the sheet was copied.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the sheet was copied.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName) =>
        workbook.CopyToFile(sheetName, new CopyToFileRequest
        {
            NewFilePath = Require(context, NewFilePath, nameof(NewFilePath)),
            NewSheetName = NewSheetName?.Get(context) ?? string.Empty,
            NewFilePassword = NewFilePassword?.Get(context) ?? string.Empty,
            NewModifyFilePassword = NewModifyFilePassword?.Get(context) ?? string.Empty,
            AutoFileCreation = AutoFileCreation?.Get(context) ?? false,
        });

    /// <inheritdoc />
    protected override void ReportResult(CodeActivityContext context, bool succeeded) =>
        ExecutionResult.Set(context, succeeded);
}
