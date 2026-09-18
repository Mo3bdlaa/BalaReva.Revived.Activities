using System.Activities;
using System.Activities.Statements;
using System.ComponentModel;

namespace BalaReva.Word;

/// <summary>
/// Opens a Word document and runs its child activities against it.
/// </summary>
/// <remarks>
/// The document is published as a workflow execution property, which is how the child
/// activities reach it, and is also handed to <see cref="Body"/> as a
/// <see cref="WordObject"/> so a workflow can bind it to a variable.
/// </remarks>
[DisplayName("Word Scope")]
[Description("Opens a Word document and runs the contained activities against it.")]
public sealed class WordScope : BaseNative
{
    /// <summary>Activities to run against the document.</summary>
    [Browsable(false)]
    public ActivityAction<WordObject> Body { get; set; } = new()
    {
        Argument = new DelegateInArgument<WordObject> { Name = "WordDocument" },
        Handler = new Sequence(),
    };

    /// <summary>
    /// The open document, kept so the scope can close it when the body finishes.
    /// </summary>
    /// <remarks>
    /// Holds the document rather than the <see cref="WordScopeHandle"/> wrapping it:
    /// WF rejects a <see cref="Variable{T}"/> whose T derives from
    /// <see cref="System.Activities.Handle"/>.
    /// </remarks>
    private Variable<IWordDocument> Opened { get; } = new();

    /// <inheritdoc />
    protected override void CacheMetadata(NativeActivityMetadata metadata)
    {
        base.CacheMetadata(metadata);
        metadata.AddImplementationVariable(Opened);
        if (Body is null) metadata.AddValidationError("Word Scope requires a body.");
    }

    /// <inheritdoc />
    protected override void Execute(NativeActivityContext context)
    {
        var path = FilePath.Get(context);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("File Path is required.", nameof(FilePath));

        var target = new WordObject
        {
            FilePath = path,
            Password = OpenPassword?.Get(context) ?? string.Empty,
            ModiPassword = ModifyPassword?.Get(context) ?? string.Empty,
        };

        var document = Service(context).Open(target.FilePath, target.Password, target.ModiPassword);
        Opened.Set(context, document);

        var handle = new WordScopeHandle { Document = document };
        context.Properties.Add(handle.ExecutionPropertyName, handle);

        if (Body is not null)
            context.ScheduleAction(Body, target, OnBodyComplete, OnBodyFault);
    }

    /// <summary>Closes the document once the body has finished.</summary>
    private void OnBodyComplete(NativeActivityContext context, ActivityInstance instance) =>
        Close(context);

    /// <summary>
    /// Closes the document when the body faults, then lets the fault carry on.
    /// </summary>
    /// <remarks>
    /// Without this a failing child would leave the document open and, with the COM
    /// implementation, a Word process running.
    /// </remarks>
    private void OnBodyFault(
        NativeActivityFaultContext context, Exception exception, ActivityInstance source) =>
        Close(context);

    private void Close(ActivityContext context)
    {
        var document = Opened.Get(context);
        if (document is null) return;

        Opened.Set(context, null!);
        document.Dispose();
    }
}
