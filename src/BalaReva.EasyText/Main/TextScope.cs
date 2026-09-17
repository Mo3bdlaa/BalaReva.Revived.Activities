using System.Activities;
using System.Activities.Statements;
using System.ComponentModel;
using BalaReva.EasyText.Base;

namespace BalaReva.EasyText.Main;

/// <summary>
/// Opens a text file and runs its child activities against it.
/// </summary>
/// <remarks>
/// The file path is published as a workflow execution property, which is how the
/// child activities reach it, and is also handed to <see cref="Body"/> as its
/// delegate argument so a workflow can bind it to a variable.
/// </remarks>
[DisplayName("Text Scope")]
[Description("Opens a text file and runs the contained activities against it.")]
public sealed class TextScope : NativeActivity
{
    /// <summary>Activities to run against the file. Receives the file path.</summary>
    [Browsable(false)]
    public ActivityAction<string> Body { get; set; } = new()
    {
        Argument = new DelegateInArgument<string> { Name = "FilePath" },
        Handler = new Sequence(),
    };

    /// <summary>Path of the text file to operate on.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("File Path")]
    [Description("Full path of the text file the contained activities operate on.")]
    public InArgument<string> FilePath { get; set; } = null!;

    /// <inheritdoc />
    protected override void CacheMetadata(NativeActivityMetadata metadata)
    {
        base.CacheMetadata(metadata);
        if (Body is null)
            metadata.AddValidationError("Text Scope requires a body.");
    }

    /// <inheritdoc />
    protected override void Execute(NativeActivityContext context)
    {
        var path = FilePath.Get(context);
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("File Path is required.", nameof(FilePath));

        // Registering under the handle's own execution property name is what lets
        // descendants resolve it with GetProperty<TextScopeHandle>().
        var handle = new TextScopeHandle { Document = new TextDocument(path) };
        context.Properties.Add(handle.ExecutionPropertyName, handle);

        if (Body is not null)
            context.ScheduleAction(Body, path);
    }
}
