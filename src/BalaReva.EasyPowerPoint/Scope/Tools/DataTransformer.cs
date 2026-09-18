using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Tools;

/// <summary>Replaces placeholder text across slides from a dictionary.</summary>
[DisplayName("Data Transformer")]
[Description("Replaces placeholder text across slides from a dictionary.")]
public sealed class DataTransformer : BaseNativeChild
{
    /// <summary>Slides to work on, numbered from 1. Empty means every slide.</summary>
    [Category("Input")]
    [DisplayName("Slide Indexes")]
    [Description("Slides to work on, numbered from 1. Empty means every slide.")]
    public InArgument<int[]> SlideIndexes { get; set; } = null!;

    /// <summary>Each key is the text to find and its value the replacement.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Input Dictionary")]
    [Description("Each key is the text to find and its value the replacement.")]
    public InArgument<Dictionary<string, string>> InputDictionary { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var replacements = InputDictionary?.Get(context)
            ?? throw new ArgumentException("InputDictionary is required.", nameof(InputDictionary));

        presentation.DataTransformer(SlideIndexes?.Get(context) ?? [], replacements);
    }
}
