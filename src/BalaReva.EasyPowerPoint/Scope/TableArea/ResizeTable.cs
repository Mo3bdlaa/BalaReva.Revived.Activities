using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Moves and resizes a table.</summary>
[DisplayName("Resize Table")]
[Description("Moves and resizes a table.")]
public sealed class ResizeTable : BaseTableNativeChild
{
    /// <summary>Distance from the left edge of the slide, in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Table Left")]
    [Description("Distance from the left edge of the slide, in points. Zero leaves it alone.")]
    public InArgument<double> TableLeft { get; set; } = null!;

    /// <summary>Distance from the top edge of the slide, in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Table Top")]
    [Description("Distance from the top edge of the slide, in points. Zero leaves it alone.")]
    public InArgument<double> TableTop { get; set; } = null!;

    /// <summary>Width in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Table Width")]
    [Description("Width in points. Zero leaves it alone.")]
    public InArgument<double> TableWidth { get; set; } = null!;

    /// <summary>Height in points. Zero leaves it alone.</summary>
    [Category("Input")]
    [DisplayName("Table Height")]
    [Description("Height in points. Zero leaves it alone.")]
    public InArgument<double> TableHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.ResizeTable(
            Table(context),
            TableLeft?.Get(context) ?? 0,
            TableTop?.Get(context) ?? 0,
            TableWidth?.Get(context) ?? 0,
            TableHeight?.Get(context) ?? 0);
}
