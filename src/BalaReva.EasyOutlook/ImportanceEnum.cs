namespace BalaReva.EasyOutlook;

/// <summary>Importance flag on an appointment or meeting.</summary>
/// <remarks>Values match <c>Microsoft.Office.Interop.Outlook.OlImportance</c>.</remarks>
public enum ImportanceEnum
{
    /// <summary>Low importance.</summary>
    Low = 0,
    /// <summary>Normal importance.</summary>
    Normal = 1,
    /// <summary>High importance.</summary>
    High = 2,
}
