using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.TableArea;

/// <summary>Adds a table to a slide, filled from a DataTable.</summary>
[DisplayName("Add Table")]
[Description("Adds a table to a slide, filled from a DataTable.")]
public sealed class AddTable : BaseSlideNativeChild
{
    /// <summary>Data to fill the table with.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Input Table")]
    [Description("Data to fill the table with.")]
    public InArgument<DataTable> InputTable { get; set; } = null!;

    /// <summary>Name to give the table shape, so later activities can find it by name.</summary>
    [Category("Input")]
    [DisplayName("Table Name")]
    [Description("Name to give the table shape, so later activities can find it by name.")]
    public InArgument<string> TableName { get; set; } = null!;

    /// <summary>Write the column names as the first row.</summary>
    [Category("Input")]
    [DisplayName("Add Header")]
    [Description("Write the column names as the first row.")]
    public bool AddHeader { get; set; }

    /// <summary>Emphasise the first column.</summary>
    [Category("Input")]
    [DisplayName("First Column")]
    [Description("Emphasise the first column.")]
    public bool FirstColumn { get; set; }

    /// <summary>Emphasise the last column.</summary>
    [Category("Input")]
    [DisplayName("Last Column")]
    [Description("Emphasise the last column.")]
    public bool LastColumn { get; set; }

    /// <summary>Distance from the left edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Table Left")]
    [Description("Distance from the left edge, in points.")]
    public InArgument<float> TableLeft { get; set; } = null!;

    /// <summary>Distance from the top edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Table Top")]
    [Description("Distance from the top edge, in points.")]
    public InArgument<float> TableTop { get; set; } = null!;

    /// <summary>Width in points. Zero lets PowerPoint choose.</summary>
    [Category("Input")]
    [DisplayName("Table Width")]
    [Description("Width in points. Zero lets PowerPoint choose.")]
    public InArgument<float> TableWidth { get; set; } = null!;

    /// <summary>Height in points. Zero lets PowerPoint choose.</summary>
    [Category("Input")]
    [DisplayName("Table Height")]
    [Description("Height in points. Zero lets PowerPoint choose.")]
    public InArgument<float> TableHeight { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var input = InputTable?.Get(context)
            ?? throw new ArgumentException("InputTable is required.", nameof(InputTable));

        presentation.AddTable(SlideIndex.Get(context), new AddTableRequest
        {
            InputTable = input,
            TableName = TableName?.Get(context) ?? string.Empty,
            AddHeader = AddHeader,
            FirstColumn = FirstColumn,
            LastColumn = LastColumn,
            Left = TableLeft?.Get(context) ?? 0,
            Top = TableTop?.Get(context) ?? 0,
            Width = TableWidth?.Get(context) ?? 0,
            Height = TableHeight?.Get(context) ?? 0,
        });
    }
}
