using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Deletes a table from a slide.</summary>
[DisplayName("Delete Table")]
[Description("Deletes a table from a slide.")]
public sealed class DeleteTable : BaseTableNativeChild
{

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.DeleteTable(Table(context));
}
