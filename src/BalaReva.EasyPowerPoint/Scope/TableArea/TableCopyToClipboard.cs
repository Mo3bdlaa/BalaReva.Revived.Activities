using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Copies a table to the Windows clipboard.</summary>
[DisplayName("Table Copy To Clipboard")]
[Description("Copies a table to the Windows clipboard.")]
public sealed class TableCopyToClipboard : BaseTableNativeChild
{

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.TableCopyToClipboard(Table(context));
}
