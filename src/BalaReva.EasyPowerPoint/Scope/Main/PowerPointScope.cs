using System.Activities;
using System.Activities.Statements;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Main;

/// <summary>
/// Opens a PowerPoint presentation and runs its child activities against it.
/// </summary>
[DisplayName("PowerPoint Scope")]
[Description("Opens a PowerPoint presentation and runs the contained activities against it.")]
public sealed class PowerPointScope : BaseNative
{
    /// <summary>Activities to run against the presentation.</summary>
    [Browsable(false)]
    public ActivityAction<PowerPointObject> Body { get; set; } = new()
    {
        Argument = new DelegateInArgument<PowerPointObject> { Name = "PowerPointPresentation" },
        Handler = new Sequence(),
    };

    /// <summary>
    /// The open presentation, kept so the scope can close it when the body finishes.
    /// </summary>
    /// <remarks>
    /// Holds the presentation rather than the handle wrapping it: WF rejects a
    /// <see cref="Variable{T}"/> whose T derives from <see cref="Handle"/>.
    /// </remarks>
    private Variable<IPowerPointPresentation> Opened { get; } = new();

    /// <inheritdoc />
    protected override void CacheMetadata(NativeActivityMetadata metadata)
    {
        base.CacheMetadata(metadata);
        metadata.AddImplementationVariable(Opened);
        if (Body is null) metadata.AddValidationError("PowerPoint Scope requires a body.");
    }

    /// <inheritdoc />
    protected override void Execute(NativeActivityContext context)
    {
        var path = FilePath.Get(context);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("File Path is required.", nameof(FilePath));

        var service = context.GetExtension<IPowerPointService>() ?? PowerPointService.Instance;
        var presentation = service.Open(
            path,
            OpenPassword?.Get(context) ?? string.Empty,
            ModifyPassword?.Get(context) ?? string.Empty,
            DisplayAlerts?.Get(context) ?? false,
            MacroSettings == EnableDisableEnum.Enable);

        Opened.Set(context, presentation);
        var handle = new PowerPointScopeHandle { Presentation = presentation };
        context.Properties.Add(handle.ExecutionPropertyName, handle);

        var target = new PowerPointObject
        {
            FilePath = path,
            Password = OpenPassword?.Get(context) ?? string.Empty,
            ModiPassword = ModifyPassword?.Get(context) ?? string.Empty,
            // Null under a stand-in service.
            PptPersentation = presentation.ComPresentation,
        };

        if (Body is not null)
            context.ScheduleAction(Body, target, OnBodyComplete, OnBodyFault);
    }

    private void OnBodyComplete(NativeActivityContext context, ActivityInstance instance) =>
        Close(context);

    /// <summary>
    /// Closes the presentation when the body faults, then lets the fault carry on.
    /// </summary>
    /// <remarks>
    /// Without this a failing child would leave a PowerPoint process running.
    /// </remarks>
    private void OnBodyFault(
        NativeActivityFaultContext context, Exception exception, ActivityInstance source) =>
        Close(context);

    private void Close(ActivityContext context)
    {
        var presentation = Opened.Get(context);
        if (presentation is null) return;
        Opened.Set(context, null!);
        presentation.Dispose();
    }
}
