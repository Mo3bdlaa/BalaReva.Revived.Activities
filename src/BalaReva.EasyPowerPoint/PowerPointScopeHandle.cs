using System.Activities;

namespace BalaReva.EasyPowerPoint;

/// <summary>
/// The execution property a <c>PowerPointScope</c> publishes so its children can find
/// the open presentation.
/// </summary>
public sealed class PowerPointScopeHandle : Handle
{
    /// <summary>The presentation the enclosing scope has open.</summary>
    public IPowerPointPresentation? Presentation { get; set; }
}
