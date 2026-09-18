using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Prints the presentation.</summary>
[DisplayName("Print")]
[Description("Prints the presentation.")]
public sealed class Print : BaseNativeChild
{
    /// <summary>How many copies to print. Less than one is treated as one.</summary>
    [Category("Input")]
    [DisplayName("Number Of Copies")]
    [Description("How many copies to print. Less than one is treated as one.")]
    public InArgument<int> NumberOfCopies { get; set; } = null!;

    /// <summary>Colour, black and white, or pure black and white.</summary>
    [Category("Input")]
    [DisplayName("Print Color Type")]
    [Description("Colour, black and white, or pure black and white.")]
    public PrintColorTypeEnum PrintColorType { get; set; } = PrintColorTypeEnum.PrintColor;

    /// <summary>Print the comments too.</summary>
    [Category("Input")]
    [DisplayName("Print Comments")]
    [Description("Print the comments too.")]
    public InArgument<bool> PrintComments { get; set; } = null!;

    /// <summary>Print hidden slides too.</summary>
    [Category("Input")]
    [DisplayName("Print Hidden Slides")]
    [Description("Print hidden slides too.")]
    public InArgument<bool> PrintHiddenSlides { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.Print(
            NumberOfCopies?.Get(context) ?? 1,
            PrintColorType,
            PrintComments?.Get(context) ?? false,
            PrintHiddenSlides?.Get(context) ?? false);
}
