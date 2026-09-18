using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Updates the presentation linked objects.</summary>
[DisplayName("Update Links")]
[Description("Updates the presentation linked objects.")]
public sealed class UpdateLinks : BaseNativeChild
{

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.UpdateLinks();
}
