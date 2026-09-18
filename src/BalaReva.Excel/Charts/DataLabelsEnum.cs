namespace BalaReva.Excel.Charts;

/// <summary>
/// Mirrors the <c>BalaReva.Excel.Charts.DataLabelsEnum</c> of the published BalaReva.Excel.Activities.
/// </summary>
/// <remarks>
/// Names and values are generated from the published assembly's metadata by
/// tools/generate_enums.py. A workflow persists the member name, and the
/// underlying number reaches .xaml through some expression forms, so neither
/// may be renamed or renumbered.
/// </remarks>
public enum DataLabelsEnum
{
    /// <summary>Show none.</summary>
    ShowNone = -4142,

    /// <summary>Show value.</summary>
    ShowValue = 2,

    /// <summary>Show percent.</summary>
    ShowPercent = 3,

    /// <summary>Show label.</summary>
    ShowLabel = 4,

    /// <summary>Show label and percent.</summary>
    ShowLabelAndPercent = 5,

    /// <summary>Show bubble sizes.</summary>
    ShowBubbleSizes = 6,
}
