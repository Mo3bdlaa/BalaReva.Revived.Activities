using System.Activities;
using System.ComponentModel;
using System.Data;
using BalaReva.EasyPowerPoint.Base;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Reads a slide comments.</summary>
[DisplayName("Comments Read")]
[Description("Reads a slide comments.")]
public sealed class CommentsRead : BaseNativeChild
{
    /// <summary>Which slide. Slides are numbered from 1.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Slide Index")]
    [Description("Which slide. Slides are numbered from 1.")]
    public InArgument<int> SlideIndex { get; set; } = null!;

    /// <summary>Include replies to each comment.</summary>
    [Category("Input")]
    [DisplayName("Include Replies")]
    [Description("Include replies to each comment.")]
    public InArgument<bool> IncludeReplies { get; set; } = null!;

    /// <summary>The comments, one per line.</summary>
    [Category("Output")]
    [DisplayName("Result")]
    [Description("The comments, one per line.")]
    public OutArgument<string> Result { get; set; } = null!;

    /// <summary>One row per comment, with its author and whether it is a reply.</summary>
    [Category("Output")]
    [DisplayName("Result Table")]
    [Description("One row per comment, with its author and whether it is a reply.")]
    public OutArgument<DataTable> ResultTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
    {
        var (text, table) = presentation.CommentsRead(
            SlideIndex.Get(context), IncludeReplies?.Get(context) ?? false);

        Result.Set(context, text);
        ResultTable.Set(context, table);
    }
}
