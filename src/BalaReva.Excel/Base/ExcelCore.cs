using System.Activities;
using System.ComponentModel;

namespace BalaReva.Excel.Base;

/// <summary>
/// Shared behaviour for the activities that open a workbook and work on one sheet.
/// </summary>
/// <remarks>
/// This package has no scope activity: each activity carries its own file path, sheet
/// name and passwords, so the base opens the workbook, runs the work and closes it
/// again. That is how the published package was shaped.
///
/// <c>ExecutionResult</c> is declared on each concrete activity rather than here, which
/// is also how the published package had it, so the base cannot set it directly.
/// <see cref="ReportResult"/> is the hook each activity overrides.
/// </remarks>
public abstract class ExcelCore : CodeActivity
{
    /// <summary>When true, a failure is reported rather than faulting the workflow.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Continue On Error")]
    [Description("Report failures through ExecutionResult instead of faulting the workflow.")]
    public InArgument<bool> ContinueOnError { get; set; } = null!;

    /// <summary>Milliseconds to wait before running.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<short> Delay { get; set; } = null!;

    /// <summary>Workbook to work on.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the workbook.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Password needed to open the workbook.</summary>
    [Category("Input")]
    [DisplayName("File Password")]
    [Description("Password needed to open the workbook, if it has one.")]
    public InArgument<string> FilePassword { get; set; } = null!;

    /// <summary>Password needed to change the workbook.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change the workbook, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;

    /// <summary>Sheet to work on.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to work on. Empty uses the workbook's active sheet.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(delay);

        try
        {
            var service = context.GetExtension<IExcelService>() ?? ExcelService.Instance;
            var path = FilePath?.Get(context);
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("FilePath is required.", nameof(FilePath));

            using var workbook = service.Open(
                path,
                FilePassword?.Get(context) ?? string.Empty,
                ModifyPassword?.Get(context) ?? string.Empty);

            ExecuteWork(context, workbook, SheetName?.Get(context) ?? string.Empty);
            ReportResult(context, true);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            ReportResult(context, false);
        }
    }

    /// <summary>Does the actual work against the open workbook.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook, string sheetName);

    /// <summary>
    /// Sets this activity's own ExecutionResult. Activities that report something else,
    /// such as <c>GetComment</c>, leave it alone.
    /// </summary>
    protected virtual void ReportResult(CodeActivityContext context, bool succeeded)
    {
    }

    /// <summary>Reads a required string argument.</summary>
    protected static string Require(CodeActivityContext context, InArgument<string> argument, string name)
    {
        var value = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required.", name);
        return value;
    }
}

/// <summary>
/// Shared behaviour for activities that do not open a workbook at all.
/// </summary>
public abstract class BaseExcel : CodeActivity
{
}

/// <summary>
/// Shared arguments for workbook-level activities that name no sheet.
/// </summary>
public abstract class BaseExcelWorkBook : CodeActivity
{
    /// <summary>Milliseconds to wait before running.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<short> Delay { get; set; } = null!;

    /// <summary>Workbook to work on.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the workbook.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Password needed to open the workbook.</summary>
    [Category("Input")]
    [DisplayName("File Password")]
    [Description("Password needed to open the workbook, if it has one.")]
    public InArgument<string> FilePassword { get; set; } = null!;

    /// <summary>Password needed to change the workbook.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change the workbook, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;
}
