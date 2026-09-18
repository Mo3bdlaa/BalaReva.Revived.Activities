using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Clears the text from every cell of a table.</summary>
[DisplayName("Clear Table")]
[Description("Clears the text from every cell of a table.")]
public sealed class ClearTable : BaseTableNativeChild
{
    /// <summary>Leave the first row alone, keeping it as a header.</summary>
    [Category("Input")]
    [DisplayName("Leave First Row")]
    [Description("Leave the first row alone, keeping it as a header.")]
    public bool LeaveFirstRow { get; set; }

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.ClearTable(Table(context), LeaveFirstRow);
}
