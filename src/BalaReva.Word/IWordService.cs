using System.Data;
using BalaReva.Word.Utilities;

namespace BalaReva.Word;

/// <summary>
/// Opening documents and the operations that do not need one open.
/// </summary>
/// <remarks>
/// The activities talk to this rather than to Word's COM object model, so that argument
/// handling and output mapping can be tested with a stand-in. As with EasyOutlook, no
/// build agent has Word installed, so <see cref="WordService"/> itself is not covered by
/// any automated test; keeping it a mechanical translation is the only defence it gets.
///
/// A workflow can register its own through <c>WorkflowInvoker.Extensions</c>.
/// </remarks>
public interface IWordService
{
    /// <summary>Opens a document for a scope to work on.</summary>
    IWordDocument Open(string filePath, string openPassword, string modifyPassword);

    /// <summary>Appends <paramref name="sourceFiles"/> onto <paramref name="wordFile"/>.</summary>
    void MergeDocuments(string wordFile, string openPassword, string modifyPassword, string[] sourceFiles);

    /// <summary>Exports a page range to PDF. Zero for either bound means "all pages".</summary>
    void WordToPdf(string wordFile, string openPassword, string modifyPassword,
                   string pdfFile, int startPage, int endPage);
}

/// <summary>
/// A document a <c>WordScope</c> has open, and everything its children can do to it.
/// </summary>
public interface IWordDocument : IDisposable
{
    /// <summary>The document the scope opened.</summary>
    WordObject Target { get; }

    // documents

    /// <summary>Sets new open and modify passwords.</summary>
    void ChangePassword(string newOpenPassword, string newModifyPassword);

    /// <summary>Creates a new, empty document at the scope's path.</summary>
    void CreateDocument();

    /// <summary>Saves a copy in another name and format.</summary>
    void SaveAs(string newFileName, EnumSaveAs format);

    /// <summary>Prints the document.</summary>
    void Print(string printerName, int copies, EnumOrientation orientation);

    // content

    /// <summary>Moves the selection to a bookmark, page or line.</summary>
    void Select(EnumGoTo target, string bookmark, int pageOrLine, int rows, int columns);

    /// <summary>Pastes the clipboard at a bookmark, page or line.</summary>
    void Paste(EnumGoTo target, string bookmark, int pageOrLine);

    /// <summary>Inserts a page break at the selection.</summary>
    void AddPageBreak();

    /// <summary>Turns off displayed line numbers.</summary>
    void RemoveDisplayLineNumber();

    /// <summary>Finds and replaces text.</summary>
    void FindReplace(string findText, string replaceText, EnumFindReplace findOption,
                     EnumReplaceOption replaceOption, bool matchCase);

    /// <summary>Runs a VBA macro and returns its result.</summary>
    object? ExecuteMacro(string macroName, object[]? parameters);

    /// <summary>Word's own counts for the document.</summary>
    WordStatisticsResult Statistics();

    /// <summary>Saves every inline image to a folder.</summary>
    void ExtractImages(string imageFolder, EnumFileExtension extension);

    // headers and footers

    /// <summary>Reads the header or footer text of every section.</summary>
    string[] ReadHeaderFooter(EnumHeadersFooters readType);

    /// <summary>Writes header or footer text.</summary>
    void InsertHeaderFooter(EnumHeadersFooters insertType, HeaderFooterText text);

    /// <summary>Places images in the header or footer.</summary>
    void InsertHeaderFooterImage(EnumHeadersFooters insertType, HeaderFooterImages images);

    /// <summary>Clears the header or footer of every section.</summary>
    void RemoveHeaderFooter(EnumHeadersFooters removeType);

    /// <summary>Saves header or footer images to a folder.</summary>
    void ExtractHeaderFooterImages(EnumHeadersFooters readType, string saveFolder);

    // readers

    /// <summary>Paragraphs carrying a font style, as an array and as a table.</summary>
    (string[] Array, DataTable Table) ReadByFont(EnumBoldItalicUnderline fontStyle);

    /// <summary>Paragraphs carrying a paragraph style, as an array and as a table.</summary>
    (string[] Array, DataTable Table) ReadByStyle(string paragraphStyle);

    // tables

    /// <summary>Number of tables in the document.</summary>
    int TableCount();

    /// <summary>Shape and style flags of one table. Tables are numbered from 1.</summary>
    WordTableInfo TableInfo(int tableIndex);

    /// <summary>Every table, one <see cref="DataTable"/> each.</summary>
    DataSet ReadAllTables(bool withHeader);

    /// <summary>Inserts an empty table at a bookmark, page or line.</summary>
    void InsertTable(EnumGoTo target, string bookmark, int pageOrLine, int rows, int columns);

    /// <summary>Inserts a filled table at a bookmark, page or line.</summary>
    void InsertDataTable(EnumGoTo target, string bookmark, int pageOrLine,
                         DataTable input, bool addHeader, string styleName);

    /// <summary>Adds rows to a table.</summary>
    void InsertTableRows(int tableIndex, int position, int count, float height, EnumInsertOption option);

    /// <summary>Adds columns to a table.</summary>
    void InsertTableColumns(int tableIndex, int position, int count, float width, EnumInsertOption option);

    /// <summary>Deletes a row. Rows are numbered from 1.</summary>
    void DeleteRow(int tableIndex, int rowIndex);

    /// <summary>Deletes a column. Columns are numbered from 1.</summary>
    void DeleteColumn(int tableIndex, int columnIndex);

    /// <summary>Deletes a whole table.</summary>
    void DeleteTable(int tableIndex);

    /// <summary>Sets the height rule of the given rows.</summary>
    void RowHeight(int tableIndex, int[] rowIndexes, float height, EnumRowHeightRule rule);

    /// <summary>Writes a cell's text and font.</summary>
    void SetTableValue(int tableIndex, int rowIndex, int columnIndex, WordCellFormat format);

    /// <summary>Applies an autofit behaviour to a table.</summary>
    void TableAutoFit(int tableIndex, EnumAutoFitBehavior autoFit);

    /// <summary>Applies a named style and its banding options to a table.</summary>
    void TableStyle(int tableIndex, string styleName, WordTableStyleOptions options);

    /// <summary>Copies a table to the clipboard.</summary>
    void CopyTableToClipboard(int tableIndex);

    // tools

    /// <summary>Closes every open Word document and quits the application.</summary>
    void CloseAllWord();
}

/// <summary>Cell text and font for <c>SetTableValue</c>.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class WordCellFormat
{
    /// <summary>Text to write into the cell.</summary>
    public string TextValue { get; set; } = string.Empty;

    /// <summary>Font name. Empty leaves it alone.</summary>
    public string FontName { get; set; } = string.Empty;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    public float FontSize { get; set; }

    /// <summary>Bold.</summary>
    public EnumSelectBoolean Bold { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Italic.</summary>
    public EnumSelectBoolean Italic { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Underline.</summary>
    public EnumSelectBoolean Underline { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Strikethrough.</summary>
    public EnumSelectBoolean Strikeout { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Superscript.</summary>
    public EnumSelectBoolean Superscript { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Subscript.</summary>
    public EnumSelectBoolean Subscript { get; set; } = EnumSelectBoolean.Select;

    /// <summary>Vertical alignment within the cell.</summary>
    public EnumCellVerticalAlignment VerticalAlignment { get; set; } = EnumCellVerticalAlignment.Select;
}

/// <summary>Banding options for <c>TableStyle</c>.</summary>
/// <remarks>Not part of the published surface; it keeps the service signature readable.</remarks>
public sealed class WordTableStyleOptions
{
    /// <summary>Style the first row as a header.</summary>
    public bool HeaderRow { get; set; }

    /// <summary>Style the last row as a totals row.</summary>
    public bool TotalRow { get; set; }

    /// <summary>Style the first column.</summary>
    public bool FirstColumn { get; set; }

    /// <summary>Style the last column.</summary>
    public bool LastColumn { get; set; }

    /// <summary>Band the rows.</summary>
    public bool BandedRows { get; set; }

    /// <summary>Band the columns.</summary>
    public bool BandedColumns { get; set; }
}
