namespace BalaReva.EasyText.Utilities;

/// <summary>
/// How the find activities compare text. Mirrors <see cref="StringComparison"/>,
/// with a <see cref="None"/> sentinel for "not chosen".
/// </summary>
/// <remarks>
/// The member names and numeric values match the published package exactly. A
/// workflow persists an enum by name, and Studio persists the underlying value in
/// some expression forms, so neither may be renamed or renumbered.
/// </remarks>
public enum StringComparisonEnum
{
    /// <summary>Not specified; the activity falls back to <see cref="Ordinal"/>.</summary>
    None = -1,
    /// <summary>Culture-sensitive comparison using the current culture.</summary>
    CurrentCulture = 0,
    /// <summary>Case-insensitive comparison using the current culture.</summary>
    CurrentCultureIgnoreCase = 1,
    /// <summary>Culture-sensitive comparison using the invariant culture.</summary>
    InvariantCulture = 2,
    /// <summary>Case-insensitive comparison using the invariant culture.</summary>
    InvariantCultureIgnoreCase = 3,
    /// <summary>Ordinal (byte-wise) comparison.</summary>
    Ordinal = 4,
    /// <summary>Case-insensitive ordinal comparison.</summary>
    OrdinalIgnoreCase = 5,
}

/// <summary>Conversion helpers for <see cref="StringComparisonEnum"/>.</summary>
public static class StringComparisonEnumExtensions
{
    /// <summary>
    /// Maps to the BCL enum, treating <see cref="StringComparisonEnum.None"/> as
    /// <see cref="StringComparison.Ordinal"/>.
    /// </summary>
    public static StringComparison ToStringComparison(this StringComparisonEnum value) =>
        value == StringComparisonEnum.None ? StringComparison.Ordinal : (StringComparison)value;
}
