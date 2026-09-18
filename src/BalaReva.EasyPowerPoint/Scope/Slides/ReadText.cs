using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Reads the text of the given slides.</summary>
[DisplayName("Read Text")]
[Description("Reads the text of the given slides.")]
public sealed class ReadText : BaseNativeChild
{
    /// <summary>Slides to read, numbered from 1. Empty means every slide.</summary>
    [Category("Input")]
    [DisplayName("Slide Indexes")]
    [Description("Slides to read, numbered from 1. Empty means every slide.")]
    public InArgument<int[]> SlideIndexes { get; set; } = null!;

    /// <summary>Prefix each line with the slide it came from.</summary>
    [Category("Input")]
    [DisplayName("Add Slide Index")]
    [Description("Prefix each line with the slide it came from.")]
    public InArgument<bool> AddSlideIndex { get; set; } = null!;

    /// <summary>Leave blank lines out.</summary>
    [Category("Input")]
    [DisplayName("Omit Empty Line")]
    [Description("Leave blank lines out.")]
    public InArgument<bool> OmitEmptyLine { get; set; } = null!;

    /// <summary>One entry per line of text.</summary>
    [Category("Output")]
    [DisplayName("Result Array")]
    [Description("One entry per line of text.")]
    public OutArgument<string[]> ResultArray { get; set; } = null!;

    /// <summary>The same lines, joined.</summary>
    [Category("Output")]
    [DisplayName("Result String")]
    [Description("The same lines, joined.")]
    public OutArgument<string> ResultString { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var (array, text) = presentation.ReadText(
            SlideIndexes?.Get(context) ?? [],
            AddSlideIndex?.Get(context) ?? false,
            OmitEmptyLine?.Get(context) ?? false);

        ResultArray.Set(context, array);
        ResultString.Set(context, text);
    }
}
