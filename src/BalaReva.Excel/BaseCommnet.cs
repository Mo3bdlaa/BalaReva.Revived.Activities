using System.Activities;
using System.ComponentModel;
using BalaReva.Excel.Base;

namespace BalaReva.Excel;

/// <summary>Shared arguments for the comment activities.</summary>
/// <remarks>
/// The class name is misspelled, and it is misspelled in the published package too.
/// It is part of the binding surface, so it stays as it is.
/// </remarks>
public abstract class BaseCommnet : ExcelCore
{
    /// <summary>Cell the comment belongs to.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Cell")]
    [Description("Cell the comment belongs to, for example B4.")]
    public InArgument<string> Cell { get; set; } = null!;
}
