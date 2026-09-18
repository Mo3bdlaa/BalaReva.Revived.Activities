using Interop = Microsoft.Office.Interop.PowerPoint;

namespace BalaReva.EasyPowerPoint.Scope.Main;

/// <summary>
/// The presentation a <see cref="PowerPointScope"/> is open on, handed to its body.
/// </summary>
/// <remarks>
/// Note <c>ModiPassword</c> and <c>PptPersentation</c>: both are misspelled in the
/// published package, and a workflow binds by property name, so both stay.
/// </remarks>
public sealed class PowerPointObject
{
    /// <summary>Full path of the open presentation.</summary>
    public string FilePath { get; set; } = string.Empty;

    /// <summary>Password used to permit changes to the presentation.</summary>
    public string ModiPassword { get; set; } = string.Empty;

    /// <summary>Password used to open the presentation.</summary>
    public string Password { get; set; } = string.Empty;

    /// <summary>
    /// The underlying COM presentation, for workflows that need to reach past these
    /// activities and drive PowerPoint themselves.
    /// </summary>
    /// <remarks>
    /// Null unless the scope is running against the real <see cref="PowerPointService"/>:
    /// the service is late-bound, so this is the one place the interop type appears, and
    /// it appears because the published package put it in the binding surface.
    /// </remarks>
    public Interop.Presentation? PptPersentation { get; set; }
}
