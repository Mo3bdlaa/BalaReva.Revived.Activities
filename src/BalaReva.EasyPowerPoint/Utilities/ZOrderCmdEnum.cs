namespace BalaReva.EasyPowerPoint.Utilities;

/// <summary>
/// Mirrors the <c>BalaReva.EasyPowerPoint.Utilities.ZOrderCmdEnum</c> of the published BalaReva.EasyPowerPoint.Activities.
/// </summary>
/// <remarks>
/// Names and values are generated from the published assembly's metadata by
/// tools/generate_enums.py. A workflow persists the member name, and the
/// underlying number reaches .xaml through some expression forms, so neither
/// may be renamed or renumbered.
/// </remarks>
public enum ZOrderCmdEnum
{
    /// <summary>None.</summary>
    None = -1,

    /// <summary>Bring to front.</summary>
    BringToFront = 0,

    /// <summary>Send to back.</summary>
    SendToBack = 1,

    /// <summary>Bring forward.</summary>
    BringForward = 2,

    /// <summary>Send backward.</summary>
    SendBackward = 3,

    /// <summary>Bring in front of text.</summary>
    BringInFrontOfText = 4,

    /// <summary>Send behind text.</summary>
    SendBehindText = 5,
}
