using System.Activities;
using System.ComponentModel;
using BalaReva.EasyExcel.Main;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Base;

/// <summary>
/// Shared behaviour for every activity that runs inside an <c>ExcelScope</c>.
/// </summary>
/// <remarks>
/// Delay is a short here, where Word and EasyPowerPoint use a double. That is how the
/// published package declared it, and a workflow binds by type as well as by name.
/// </remarks>
public abstract class ExcelActivity : CodeActivity
{
    /// <summary>
    /// When true, a failure is reported through <see cref="ExecutionResult"/> instead of
    /// faulting the workflow.
    /// </summary>
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

    /// <summary>True when the activity completed without error.</summary>
    [Category("Output")]
    [DisplayName("Execution Result")]
    [Description("True when the activity completed without error.")]
    public OutArgument<bool> ExecutionResult { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(TimeSpan.FromMilliseconds(delay));

        try
        {
            ExecuteWork(context, Resolve(context));
            ExecutionResult.Set(context, true);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            ExecutionResult.Set(context, false);
        }
    }

    /// <summary>Does the actual work against the scope's open workbook.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook);

    /// <summary>Finds the workbook opened by the enclosing scope.</summary>
    private static IExcelWorkbook Resolve(CodeActivityContext context) =>
        context.GetProperty<ExcelScopeHandle>()?.Workbook
        ?? throw new InvalidOperationException(
            "This activity must be placed inside an Excel Scope activity, which opens the workbook.");

    /// <summary>Reads a required string argument.</summary>
    protected static string Require(CodeActivityContext context, InArgument<string> argument, string name)
    {
        var value = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required.", name);
        return value;
    }
}

/// <summary>An <see cref="ExcelActivity"/> that works on one sheet.</summary>
public abstract class BaseActivity : ExcelActivity
{
    /// <summary>Sheet to work on. Empty means the active sheet.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to work on. Empty means the active sheet.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void ExecuteWork(CodeActivityContext context, IExcelWorkbook workbook) =>
        ExecuteWork(context, workbook, SheetName?.Get(context) ?? string.Empty);

    /// <summary>Does the actual work against one sheet of the scope's workbook.</summary>
    protected abstract void ExecuteWork(
        CodeActivityContext context, IExcelWorkbook workbook, string sheetName);
}

/// <summary>
/// Shared behaviour for the two activities that read a workbook without opening Excel.
/// </summary>
/// <remarks>
/// These go through the Open XML package format instead, so they carry their own file
/// path rather than relying on a scope, and they work on a machine with no Excel at all.
/// </remarks>
public abstract class BaseOpenXml : CodeActivity
{
    /// <summary>When true, failures are swallowed rather than faulting the workflow.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Continue On Error")]
    [Description("Swallow failures instead of faulting the workflow.")]
    public InArgument<bool> ContinueOnError { get; set; } = null!;

    /// <summary>Milliseconds to wait before running.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<short> Delay { get; set; } = null!;

    /// <summary>Full path of the workbook to read.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the workbook to read.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Sheet to read. Empty means the first sheet.</summary>
    [Category("Input")]
    [DisplayName("Sheet Name")]
    [Description("Sheet to read. Empty means the first sheet.")]
    public InArgument<string> SheetName { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(TimeSpan.FromMilliseconds(delay));

        try
        {
            var path = FilePath?.Get(context);
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("File Path is required.", nameof(FilePath));

            ExecuteWork(
                context,
                context.GetExtension<IOpenXmlReader>() ?? OpenXmlReader.Instance,
                path,
                SheetName?.Get(context) ?? string.Empty);
        }
        catch (Exception) when (ContinueOnError.Get(context))
        {
            // Reported nowhere: this base has no ExecutionResult, which is how the
            // published package declared it.
        }
    }

    /// <summary>Does the actual work against the file.</summary>
    protected abstract void ExecuteWork(
        CodeActivityContext context, IOpenXmlReader reader, string filePath, string sheetName);
}

/// <summary>Shared arguments for the activity that opens a workbook.</summary>
public abstract class BaseScope : NativeActivity
{
    /// <summary>Let Excel show its own prompts and warnings.</summary>
    [Category("Input")]
    [DisplayName("Display Alerts")]
    [Description("Let Excel show its own prompts and warnings.")]
    public InArgument<bool> DisplayAlerts { get; set; } = null!;

    /// <summary>Password needed to open the workbook.</summary>
    [Category("Input")]
    [DisplayName("File Password")]
    [Description("Password needed to open the workbook, if it has one.")]
    public InArgument<string> FilePassword { get; set; } = null!;

    /// <summary>Full path of the workbook.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the workbook to open.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Whether macros in the workbook may run.</summary>
    [Category("Input")]
    [DisplayName("Macro Settings")]
    [Description("Whether macros in the workbook may run.")]
    public EnableDisableEnum MacroSettings { get; set; } = EnableDisableEnum.Disable;

    /// <summary>Password needed to change the workbook.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change the workbook, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;

    /// <summary>Whether links to other workbooks are refreshed on open.</summary>
    [Category("Input")]
    [DisplayName("Update Auto Links")]
    [Description("Whether links to other workbooks are refreshed when the file opens.")]
    public AutomaticLinkEnum UpdateAutoLinks { get; set; } = AutomaticLinkEnum.Default;

    /// <summary>Show the Excel window while the scope runs.</summary>
    [Category("Input")]
    [DisplayName("Visible")]
    [Description("Show the Excel window while the scope runs.")]
    public InArgument<bool> Visible { get; set; } = null!;

    /// <summary>
    /// The Excel session to talk to: a workflow extension when one is registered,
    /// otherwise the real COM implementation.
    /// </summary>
    private protected static IExcelService Service(NativeActivityContext context) =>
        context.GetExtension<IExcelService>() ?? ExcelService.Instance;
}
