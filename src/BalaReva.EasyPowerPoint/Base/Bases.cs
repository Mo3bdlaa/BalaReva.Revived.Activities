using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Base;

/// <summary>Shared arguments for the activity that opens a presentation.</summary>
public abstract class BaseNative : NativeActivity
{
    /// <summary>Full path of the presentation.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the PowerPoint presentation to open.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Password needed to open the presentation.</summary>
    [Category("Input")]
    [DisplayName("Open Password")]
    [Description("Password needed to open the presentation, if it has one.")]
    public InArgument<string> OpenPassword { get; set; } = null!;

    /// <summary>Password needed to change the presentation.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change the presentation, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;

    /// <summary>Whether PowerPoint may show dialogs while the scope runs.</summary>
    [Category("Input")]
    [DisplayName("Display Alerts")]
    [Description("Let PowerPoint show its own dialogs. Off by default, so a robot is not blocked.")]
    public InArgument<bool> DisplayAlerts { get; set; } = null!;

    /// <summary>Whether macros in the presentation are enabled.</summary>
    [Category("Input")]
    [DisplayName("Macro Settings")]
    [Description("Whether macros in the presentation are enabled.")]
    public EnableDisableEnum MacroSettings { get; set; } = EnableDisableEnum.Disable;
}

/// <summary>Shared behaviour for every activity inside a <c>PowerPointScope</c>.</summary>
/// <remarks>
/// <c>Delay</c> is a double here, as it is in the Word package and unlike the short used
/// by Excel and Printer. That is how the published package declared it.
/// </remarks>
public abstract class BaseNativeChild : CodeActivity
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
    public InArgument<double> Delay { get; set; } = null!;

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

    /// <summary>Does the actual work against the scope's open presentation.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation);

    /// <summary>Finds the presentation opened by the enclosing scope.</summary>
    private static IPowerPointPresentation Resolve(CodeActivityContext context) =>
        context.GetProperty<PowerPointScopeHandle>()?.Presentation
        ?? throw new InvalidOperationException(
            "This activity must be placed inside a PowerPoint Scope activity, "
            + "which opens the presentation.");

    /// <summary>Reads a required string argument.</summary>
    protected static string Require(CodeActivityContext context, InArgument<string> argument, string name)
    {
        var value = argument?.Get(context);
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException($"{name} is required.", name);
        return value;
    }
}

/// <summary>An in-scope activity that works on one slide.</summary>
public abstract class BaseSlideNativeChild : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide to work on. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;
}

/// <summary>An in-scope activity that works on one table on one slide.</summary>
/// <remarks>
/// A table can be addressed by index or by name; the service prefers the name when both
/// are given, since a name survives slides being reordered.
/// </remarks>
public abstract class BaseTableNativeChild : BaseSlideNativeChild
{
    /// <summary>Which table on the slide. Tables are numbered from 1.</summary>
    [Category("Input")]
    [DisplayName("Table Index")]
    [Description("Which table on the slide. Tables are numbered from 1. Ignored when Table Name is set.")]
    public InArgument<int> TableIndex { get; set; } = null!;

    /// <summary>Name of the table shape.</summary>
    [Category("Input")]
    [DisplayName("Table Name")]
    [Description("Name of the table shape. Takes precedence over Table Index.")]
    public InArgument<string> TableName { get; set; } = null!;

    /// <summary>Gathers the two ways of addressing a table.</summary>
    protected TableRef Table(CodeActivityContext context) => new()
    {
        SlideIndex = SlideIndex.Get(context),
        TableIndex = TableIndex?.Get(context) ?? 0,
        TableName = TableName?.Get(context) ?? string.Empty,
    };
}

/// <summary>Shared arguments for activities that open a presentation on their own.</summary>
public abstract class BasePowerPoint : CodeActivity
{
    /// <summary>Milliseconds to wait before running.</summary>
    [RequiredArgument]
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<double> Delay { get; set; } = null!;

    /// <summary>Full path of the presentation.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the presentation.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <summary>Password needed to open the presentation.</summary>
    [Category("Input")]
    [DisplayName("Open Password")]
    [Description("Password needed to open the presentation, if it has one.")]
    public InArgument<string> OpenPassword { get; set; } = null!;

    /// <summary>Password needed to change the presentation.</summary>
    [Category("Input")]
    [DisplayName("Modify Password")]
    [Description("Password needed to change the presentation, if it has one.")]
    public InArgument<string> ModifyPassword { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay.Get(context);
        if (delay > 0) Thread.Sleep(TimeSpan.FromMilliseconds(delay));

        var service = context.GetExtension<IPowerPointService>() ?? PowerPointService.Instance;
        var path = FilePath?.Get(context);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("FilePath is required.", nameof(FilePath));

        using var presentation = service.Open(
            path,
            OpenPassword?.Get(context) ?? string.Empty,
            ModifyPassword?.Get(context) ?? string.Empty,
            displayAlerts: false,
            macrosEnabled: true);

        ExecuteWork(context, presentation);
    }

    /// <summary>Does the actual work against the presentation this activity opened.</summary>
    protected abstract void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation);
}

/// <summary>Which table an activity means.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class TableRef
{
    /// <summary>Slide the table is on, numbered from 1.</summary>
    public int SlideIndex { get; set; }

    /// <summary>Table's position on the slide, numbered from 1.</summary>
    public int TableIndex { get; set; }

    /// <summary>Table shape's name, which wins over the index when set.</summary>
    public string TableName { get; set; } = string.Empty;
}
