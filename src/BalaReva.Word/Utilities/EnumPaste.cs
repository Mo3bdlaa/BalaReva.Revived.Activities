namespace BalaReva.Word.Utilities;

/// <summary>
/// Mirrors the <c>BalaReva.Word.Utilities.EnumPaste</c> of the published BalaReva.Word.Activities.
/// </summary>
/// <remarks>
/// Names and values are generated from the published assembly's metadata by
/// tools/generate_enums.py. A workflow persists the member name, and the
/// underlying number reaches .xaml through some expression forms, so neither
/// may be renamed or renumbered.
/// </remarks>
public enum EnumPaste
{
    /// <summary>Paste default.</summary>
    wdPasteDefault = 0,

    /// <summary>Chart picture.</summary>
    wdChartPicture = 13,

    /// <summary>Chart.</summary>
    wdChart = 14,

    /// <summary>Chart linked.</summary>
    wdChartLinked = 15,

    /// <summary>Format original formatting.</summary>
    wdFormatOriginalFormatting = 16,

    /// <summary>Use destination styles recovery.</summary>
    wdUseDestinationStylesRecovery = 19,

    /// <summary>Format surrounding formatting with emphasis.</summary>
    wdFormatSurroundingFormattingWithEmphasis = 20,

    /// <summary>Format plain text.</summary>
    wdFormatPlainText = 22,
}
