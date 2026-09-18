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
    /// <para>
    /// This is the one place the revived package cannot match the published one. There
    /// it is a <c>Microsoft.Office.Interop.PowerPoint.Presentation</c>; here it is an
    /// object, and a workflow that wants the interop type has to cast.
    /// </para>
    /// <para>
    /// Declaring it as Presentation costs more than it buys. Every Office interop
    /// assembly on NuGet has a hard reference on 'office' (Microsoft.Office.Core), which
    /// Microsoft publishes nowhere, and the CLR goes looking for it as soon as it loads
    /// a member typed that way. That is not a failure confined to this property: the
    /// scope hands this object to its body, so the whole scope stops working on any
    /// machine without office.dll. An object that needs a cast is better than a scope
    /// that will not run.
    /// </para>
    /// <para>Null unless the scope is running against the real service.</para>
    /// </remarks>
    public object? PptPersentation { get; set; }
}
