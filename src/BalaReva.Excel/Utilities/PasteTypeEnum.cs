namespace BalaReva.Excel.Utilities;

/// <summary>
/// Mirrors the <c>BalaReva.Excel.Utilities.PasteTypeEnum</c> of the published BalaReva.Excel.Activities.
/// </summary>
/// <remarks>
/// Names and values are generated from the published assembly's metadata by
/// tools/generate_enums.py. A workflow persists the member name, and the
/// underlying number reaches .xaml through some expression forms, so neither
/// may be renamed or renumbered.
/// </remarks>
public enum PasteTypeEnum
{
    /// <summary>Values.</summary>
    Values = -4163,

    /// <summary>Comments.</summary>
    Comments = -4144,

    /// <summary>Formulas.</summary>
    Formulas = -4123,

    /// <summary>Formats.</summary>
    Formats = -4122,

    /// <summary>All.</summary>
    All = -4104,

    /// <summary>Validation.</summary>
    Validation = 6,

    /// <summary>All except borders.</summary>
    AllExceptBorders = 7,

    /// <summary>Column widths.</summary>
    ColumnWidths = 8,

    /// <summary>Formulas and number formats.</summary>
    FormulasAndNumberFormats = 11,

    /// <summary>Values and number formats.</summary>
    ValuesAndNumberFormats = 12,

    /// <summary>All using source theme.</summary>
    AllUsingSourceTheme = 13,

    /// <summary>All merging conditional formats.</summary>
    AllMergingConditionalFormats = 14,
}
