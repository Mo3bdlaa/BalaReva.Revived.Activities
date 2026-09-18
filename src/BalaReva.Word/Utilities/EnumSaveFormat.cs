namespace BalaReva.Word.Utilities;

/// <summary>
/// Mirrors the <c>BalaReva.Word.Utilities.EnumSaveFormat</c> of the published BalaReva.Word.Activities.
/// </summary>
/// <remarks>
/// Names and values are generated from the published assembly's metadata by
/// tools/generate_enums.py. A workflow persists the member name, and the
/// underlying number reaches .xaml through some expression forms, so neither
/// may be renamed or renumbered.
/// </remarks>
public enum EnumSaveFormat
{
    /// <summary>Document.</summary>
    Document = 0,

    /// <summary>Template.</summary>
    Template = 1,

    /// <summary>Text.</summary>
    Text = 2,

    /// <summary>Text line breaks.</summary>
    TextLineBreaks = 3,

    /// <summary>Dos text.</summary>
    DOSText = 4,

    /// <summary>Dos text line breaks.</summary>
    DOSTextLineBreaks = 5,

    /// <summary>Rtf.</summary>
    RTF = 6,

    /// <summary>Unicode text encoded text.</summary>
    UnicodeText_EncodedText = 7,

    /// <summary>Html.</summary>
    HTML = 8,

    /// <summary>Web archive.</summary>
    WebArchive = 9,

    /// <summary>Filtered html.</summary>
    FilteredHTML = 10,

    /// <summary>Xml.</summary>
    XML = 11,

    /// <summary>Xml document.</summary>
    XMLDocument = 12,

    /// <summary>Xml document macro enabled.</summary>
    XMLDocumentMacroEnabled = 13,

    /// <summary>Xml template.</summary>
    XMLTemplate = 14,

    /// <summary>Xml template macro enabled.</summary>
    XMLTemplateMacroEnabled = 15,

    /// <summary>Document default.</summary>
    DocumentDefault = 16,

    /// <summary>Pdf.</summary>
    PDF = 17,

    /// <summary>Xps.</summary>
    XPS = 18,

    /// <summary>Flat xml.</summary>
    FlatXML = 19,

    /// <summary>Flat xml macro enabled.</summary>
    FlatXMLMacroEnabled = 20,

    /// <summary>Flat xml template.</summary>
    FlatXMLTemplate = 21,

    /// <summary>Flat xml template macro enabled.</summary>
    FlatXMLTemplateMacroEnabled = 22,

    /// <summary>Open document text.</summary>
    OpenDocumentText = 23,

    /// <summary>Strict open xml document.</summary>
    StrictOpenXMLDocument = 24,
}
