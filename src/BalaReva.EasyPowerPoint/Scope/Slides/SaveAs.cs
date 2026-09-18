using System.Activities;
using System.ComponentModel;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Saves a copy under another name and format.</summary>
[DisplayName("Save As")]
[Description("Saves a copy under another name and format.")]
public sealed class SaveAs : BaseNativeChild
{
    /// <summary>Full path to save the copy to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Save As File")]
    [Description("Full path to save the copy to.")]
    public InArgument<string> SaveAsFile { get; set; } = null!;

    /// <summary>Format to save the copy in.</summary>
    [Category("Input")]
    [DisplayName("Save As Format")]
    [Description("Format to save the copy in.")]
    public SaveAsEnum SaveAsFormat { get; set; } = SaveAsEnum.Presentation;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.SaveAs(Require(context, SaveAsFile, nameof(SaveAsFile)), SaveAsFormat);
}
