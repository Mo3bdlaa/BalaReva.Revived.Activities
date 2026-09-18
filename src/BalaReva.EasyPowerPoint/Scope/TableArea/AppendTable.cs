using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Adds a row to the end of a table.</summary>
[DisplayName("Append Table")]
[Description("Adds a row to the end of a table.")]
public sealed class AppendTable : BaseTableNativeChild
{
    /// <summary>Values for the new row, one per column. Extra values are ignored.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Array Value")]
    [Description("Values for the new row, one per column. Extra values are ignored.")]
    public InArgument<string[]> ArrayValue { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.AppendTable(Table(context), ArrayValue?.Get(context) ?? []);
}
