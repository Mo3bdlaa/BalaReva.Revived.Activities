using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Adds a comment to a slide.</summary>
[DisplayName("Comments Add")]
[Description("Adds a comment to a slide.")]
public sealed class CommentsAdd : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Who the comment is from.</summary>
    [Category("Input")]
    [DisplayName("Author")]
    [Description("Who the comment is from.")]
    public InArgument<string> Author { get; set; } = null!;

    /// <summary>Text of the comment.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Comment Text")]
    [Description("Text of the comment.")]
    public InArgument<string> CommentText { get; set; } = null!;

    /// <summary>Distance from the left edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Left")]
    [Description("Distance from the left edge, in points.")]
    public InArgument<float> Left { get; set; } = null!;

    /// <summary>Distance from the top edge, in points.</summary>
    [Category("Input")]
    [DisplayName("Top")]
    [Description("Distance from the top edge, in points.")]
    public InArgument<float> Top { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.CommentsAdd(
            SlideIndex.Get(context),
            Author?.Get(context) ?? string.Empty,
            Require(context, CommentText, nameof(CommentText)),
            Left?.Get(context) ?? 0,
            Top?.Get(context) ?? 0);
}
