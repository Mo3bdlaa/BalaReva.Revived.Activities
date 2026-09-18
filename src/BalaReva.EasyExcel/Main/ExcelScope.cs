using System.Activities;
using System.Activities.Statements;
using System.ComponentModel;
using BalaReva.EasyExcel.Base;
using BalaReva.EasyExcel.Utilities;

namespace BalaReva.EasyExcel.Main;

/// <summary>
/// Opens an Excel workbook and runs its child activities against it.
/// </summary>
/// <remarks>
/// The workbook is published as a workflow execution property, which is how the child
/// activities reach it, and is also handed to <see cref="Body"/> as an
/// <see cref="ExcelParam"/> so a workflow can bind it to a variable.
/// </remarks>
[DisplayName("Excel Scope")]
[Description("Opens an Excel workbook and runs the contained activities against it.")]
public sealed class ExcelScope : BaseScope
{
    /// <summary>Activities to run against the workbook.</summary>
    [Browsable(false)]
    public ActivityAction<ExcelParam> Body { get; set; } = new()
    {
        Argument = new DelegateInArgument<ExcelParam> { Name = "ExcelWorkBook" },
        Handler = new Sequence(),
    };

    /// <summary>
    /// The open workbook, kept so the scope can close it when the body finishes.
    /// </summary>
    /// <remarks>
    /// Holds the workbook rather than the <see cref="ExcelScopeHandle"/> wrapping it:
    /// WF rejects a <see cref="Variable{T}"/> whose T derives from
    /// <see cref="Handle"/>.
    /// </remarks>
    private Variable<IExcelWorkbook> Opened { get; } = new();

    /// <inheritdoc />
    protected override void CacheMetadata(NativeActivityMetadata metadata)
    {
        base.CacheMetadata(metadata);
        metadata.AddImplementationVariable(Opened);
        if (Body is null) metadata.AddValidationError("Excel Scope requires a body.");
    }

    /// <inheritdoc />
    protected override void Execute(NativeActivityContext context)
    {
        var path = FilePath.Get(context);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("File Path is required.", nameof(FilePath));

        var workbook = Service(context).Open(new ExcelOpenRequest
        {
            FilePath = path,
            FilePassword = FilePassword?.Get(context) ?? string.Empty,
            ModifyPassword = ModifyPassword?.Get(context) ?? string.Empty,
            DisplayAlerts = DisplayAlerts?.Get(context) ?? false,
            Visible = Visible?.Get(context) ?? false,
            MacrosEnabled = MacroSettings == EnableDisableEnum.Enable,
            UpdateAutoLinks = UpdateAutoLinks,
        });

        Opened.Set(context, workbook);

        var handle = new ExcelScopeHandle { Workbook = workbook };
        context.Properties.Add(handle.ExecutionPropertyName, handle);

        var target = new ExcelParam
        {
            FilePath = path,
            Password = FilePassword?.Get(context) ?? string.Empty,
            ModiPassword = ModifyPassword?.Get(context) ?? string.Empty,
            // Null under a stand-in service.
            ExcelWorkBook = workbook.ComWorkbook,
        };

        if (Body is not null)
            context.ScheduleAction(Body, target, OnBodyComplete, OnBodyFault);
    }

    private void OnBodyComplete(NativeActivityContext context, ActivityInstance instance) =>
        Close(context);

    /// <summary>
    /// Closes the workbook when the body faults, then lets the fault carry on.
    /// </summary>
    /// <remarks>Without this a failing child would leave an Excel process running.</remarks>
    private void OnBodyFault(
        NativeActivityFaultContext context, Exception exception, ActivityInstance source) =>
        Close(context);

    private void Close(ActivityContext context)
    {
        var workbook = Opened.Get(context);
        if (workbook is null) return;
        Opened.Set(context, null!);
        workbook.Dispose();
    }
}
