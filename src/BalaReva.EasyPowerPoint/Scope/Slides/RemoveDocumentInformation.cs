using System.Activities;
using System.ComponentModel;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Scope.Slides;

/// <summary>Strips document information of the given kind.</summary>
[DisplayName("Remove Document Information")]
[Description("Strips document information of the given kind.")]
public sealed class RemoveDocumentInformation : BaseNativeChild
{
    /// <summary>Which kind of document information to strip.</summary>
    [Category("Input")]
    [DisplayName("Doc Info Type")]
    [Description("Which kind of document information to strip.")]
    public RemoveDocInfoTypeEnum DocInfoType { get; set; } = RemoveDocInfoTypeEnum.DocumentProperties;

    /// <inheritdoc />
    protected override void ExecuteWork(CodeActivityContext context, IPowerPointPresentation presentation)
        => presentation.RemoveDocumentInformation(DocInfoType);
}
