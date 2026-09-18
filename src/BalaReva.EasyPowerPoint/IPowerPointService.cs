using System.Data;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;
using BalaReva.PowerPoint;

namespace BalaReva.EasyPowerPoint;

/// <summary>Opens presentations.</summary>
/// <remarks>
/// The activities talk to this rather than to PowerPoint's COM object model, so argument
/// handling and output mapping can be tested with a stand-in. No build agent has
/// PowerPoint installed, so <see cref="PowerPointService"/> itself is not covered by any
/// automated test.
/// </remarks>
public interface IPowerPointService
{
    /// <summary>Opens a presentation, creating it when the path does not exist.</summary>
    IPowerPointPresentation Open(string filePath, string openPassword, string modifyPassword,
                                 bool displayAlerts, bool macrosEnabled);
}

/// <summary>An open presentation and everything the activities do to it.</summary>
public interface IPowerPointPresentation : IDisposable
{
    /// <summary>Full path of the open presentation.</summary>
    string FilePath { get; }

    /// <summary>
    /// The underlying COM presentation, or null when there is not one.
    /// </summary>
    /// <remarks>
    /// Typed as object because this assembly drives PowerPoint late-bound; the scope
    /// casts it to hand a workflow the interop type the published package exposed on
    /// <c>PowerPointObject.PptPersentation</c>.
    /// </remarks>
    object? ComPresentation { get; }

    // slides

    /// <summary>Number of slides.</summary>
    int SlideCount();

    /// <summary>Adds an empty slide at the given position.</summary>
    void NewSlide(int slideIndex);

    /// <summary>Deletes a slide.</summary>
    void DeleteSlide(int slideIndex);

    /// <summary>Duplicates a slide in place.</summary>
    void DuplicateSlide(int slideIndex);

    /// <summary>Copies a slide to the clipboard.</summary>
    void SlideCopy(int slideIndex);

    /// <summary>Pastes the clipboard's slide after the given position.</summary>
    void SlidePaste(int slideIndex);

    /// <summary>Reads a slide's text shapes: their text, font and position.</summary>
    SlideObject SlideExtractor(int slideIndex);

    /// <summary>Shows or hides a slide during a slide show.</summary>
    void HideUnhideSlide(int slideIndex, HideUnhideEnum slideShow);

    /// <summary>Sets a slide's transition.</summary>
    void SlideTransitions(int slideIndex, EntryEffectEnum effect, bool onMouseClick, float duration);

    /// <summary>Appends slides from another presentation.</summary>
    void Merge(string sourcePpt, int startSlideIndex, int endSlideIndex, int slideAfter);

    // text and shapes

    /// <summary>Reads the text of the given slides.</summary>
    (string[] Array, string Text) ReadText(int[] slideIndexes, bool addSlideIndex, bool omitEmptyLine);

    /// <summary>Finds text across the given slides, reporting which slides matched.</summary>
    (int[] SlideIndexes, DataTable Table) FindText(int[] slideIndexes, string find,
                                                   bool matchCase, bool wholeWord);

    /// <summary>Replaces text across the given slides.</summary>
    void FindReplace(int[] slideIndexes, string find, string replace,
                     bool matchCase, bool wholeWord, bool firstOccurrence);

    /// <summary>Adds a text box to a slide.</summary>
    void InsertTextBox(int slideIndex, TextBoxRequest request);

    /// <summary>Restyles one text shape on a slide.</summary>
    void TextShapeEdit(int slideIndex, int textIndex, TextShape style);

    /// <summary>Number of text shapes on a slide.</summary>
    int TextShapeCount(int slideIndex);

    /// <summary>Hyperlinks found on the given slides.</summary>
    DataTable ExtractHyperLinks(int[] slideIndexes);

    // images

    /// <summary>Places an image on a slide.</summary>
    void InsertPicture(int slideIndex, PictureRequest request);

    /// <summary>Deletes an image from a slide.</summary>
    void DeleteImage(int slideIndex, int imageIndex);

    /// <summary>Number of image shapes on a slide.</summary>
    int ImageShapeCount(int slideIndex);

    /// <summary>Saves a slide's images into a folder.</summary>
    void ImageExtractor(int slideIndex, string imageDirectory, ImageFileFormatEnum format);

    /// <summary>Pastes the clipboard onto a slide at a position and size.</summary>
    void PasteClipboard(int slideIndex, double left, double top, double width, double height);

    // charts

    /// <summary>Number of chart shapes on a slide.</summary>
    int ChartShapeCount(int slideIndex);

    /// <summary>Deletes a chart from a slide.</summary>
    void ChartDelete(int slideIndex, int chartIndex);

    /// <summary>Copies a chart to the clipboard.</summary>
    void ChartCopyToClipboard(int slideIndex, int chartIndex);

    /// <summary>Moves and resizes a chart.</summary>
    void ChartFormat(int slideIndex, int chartIndex, double left, double top,
                     double width, double height);

    /// <summary>Saves a slide's charts as images.</summary>
    void ChartImageExtract(int slideIndex, string imageFolder, ImageFileFormatEnum format);

    /// <summary>Refreshes the data behind the given slides' charts.</summary>
    void RefreshData(short[] slideIndexes);

    /// <summary>Updates the presentation's linked objects.</summary>
    void UpdateLinks();

    // comments

    /// <summary>Adds a comment to a slide.</summary>
    void CommentsAdd(int slideIndex, string author, string commentText, float left, float top);

    /// <summary>Deletes every comment on a slide.</summary>
    void CommentsDelete(int slideIndex);

    /// <summary>Reads a slide's comments.</summary>
    (string[] Comments, DataTable Table) CommentsRead(int slideIndex, bool includeReplies);

    // tables

    /// <summary>Adds a table to a slide, filled from a DataTable.</summary>
    void AddTable(int slideIndex, AddTableRequest request);

    /// <summary>Deletes a table.</summary>
    void DeleteTable(TableRef table);

    /// <summary>Clears a table's cells.</summary>
    void ClearTable(TableRef table, bool leaveFirstRow);

    /// <summary>Appends a row of values to a table.</summary>
    void AppendTable(TableRef table, string[] values);

    /// <summary>Writes one cell.</summary>
    void EditTable(TableRef table, int rowIndex, int columnIndex, string cellValue);

    /// <summary>Reads one cell.</summary>
    string GetRowItem(TableRef table, int rowIndex, int columnIndex);

    /// <summary>Deletes a table row.</summary>
    void DeleteRow(TableRef table, int rowIndex);

    /// <summary>Deletes a table column.</summary>
    void DeleteColumn(TableRef table, int columnIndex);

    /// <summary>Moves and resizes a table.</summary>
    void ResizeTable(TableRef table, double left, double top, double width, double height);

    /// <summary>Sets a table's font.</summary>
    void FontOption(TableRef table, TextStyleRequest style);

    /// <summary>Sets a table's banding and emphasis options.</summary>
    void StyleOption(TableRef table, TableStyleOptions options);

    /// <summary>Copies a table to the clipboard.</summary>
    void TableCopyToClipboard(TableRef table);

    /// <summary>Every table on the given slides, one DataTable each.</summary>
    DataTable[] ExtractTables(int[] slideIndexes, bool hasHeader);

    /// <summary>Names of the table shapes on a slide.</summary>
    string[] GetTableNames(int slideIndex);

    // tools

    /// <summary>Replaces placeholder text across slides from a dictionary.</summary>
    void DataTransformer(int[] slideIndexes, Dictionary<string, string> replacements);

    // presentation

    /// <summary>Exports the presentation to PDF or XPS.</summary>
    void ExportPdf(string filePath, FixedFormatIntentEnum formatType);

    /// <summary>Saves a copy in another name and format.</summary>
    void SaveAs(string saveAsFile, SaveAsEnum format);

    /// <summary>Prints the presentation.</summary>
    void Print(int numberOfCopies, PrintColorTypeEnum colorType,
               TrueFalseNoneEnum printComments, TrueFalseNoneEnum printHiddenSlides);

    /// <summary>Strips document information of the given kind.</summary>
    void RemoveDocumentInformation(RemoveDocInfoTypeEnum docInfoType);

    /// <summary>Runs a VBA macro and returns its result.</summary>
    object? ExecuteMacro(string macroName, object[]? arguments);
}

/// <summary>A text box to add to a slide.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class TextBoxRequest
{
    /// <summary>Text to put in the box.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    public float Left { get; set; }

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    public float Top { get; set; }

    /// <summary>Width in points.</summary>
    public float Width { get; set; }

    /// <summary>Height in points.</summary>
    public float Height { get; set; }

    /// <summary>How the text sits inside the box.</summary>
    public EnumTextEffectAlignment Alignment { get; set; } = EnumTextEffectAlignment.Left;

    /// <summary>Which way the text runs.</summary>
    public TextOrientationEnum Orientation { get; set; } = TextOrientationEnum.Horizontal;

    /// <summary>Where the box sits in the slide's stacking order.</summary>
    public ZOrderCmdEnum ZOrder { get; set; } = ZOrderCmdEnum.BringToFront;

    /// <summary>Font and emphasis for the text.</summary>
    public TextStyleRequest Style { get; set; } = new();
}

/// <summary>Font and emphasis.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class TextStyleRequest
{
    /// <summary>Font name. Empty leaves it alone.</summary>
    public string FontName { get; set; } = string.Empty;

    /// <summary>Font size in points. Zero leaves it alone.</summary>
    public float FontSize { get; set; }

    /// <summary>Bold. None leaves it alone.</summary>
    public TrueFalseNoneEnum Bold { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Italic. None leaves it alone.</summary>
    public TrueFalseNoneEnum Italic { get; set; } = TrueFalseNoneEnum.None;
}

/// <summary>An image to place on a slide.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class PictureRequest
{
    /// <summary>Full path of the image.</summary>
    public string ImagePath { get; set; } = string.Empty;

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    public float Left { get; set; }

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    public float Top { get; set; }

    /// <summary>Width in points. Zero keeps the image's own width.</summary>
    public float Width { get; set; }

    /// <summary>Height in points. Zero keeps the image's own height.</summary>
    public float Height { get; set; }

    /// <summary>Where the image sits in the slide's stacking order.</summary>
    public ZOrderCmdEnum ZOrder { get; set; } = ZOrderCmdEnum.BringToFront;
}

/// <summary>A table to add to a slide.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class AddTableRequest
{
    /// <summary>Data to fill the table with.</summary>
    public DataTable InputTable { get; set; } = new();

    /// <summary>Name to give the table shape.</summary>
    public string TableName { get; set; } = string.Empty;

    /// <summary>Write the column names as the first row.</summary>
    public bool AddHeader { get; set; }

    /// <summary>Emphasise the first column.</summary>
    public bool FirstColumn { get; set; }

    /// <summary>Emphasise the last column.</summary>
    public bool LastColumn { get; set; }

    /// <summary>Distance from the left edge of the slide, in points.</summary>
    public float Left { get; set; }

    /// <summary>Distance from the top edge of the slide, in points.</summary>
    public float Top { get; set; }

    /// <summary>Width in points.</summary>
    public float Width { get; set; }

    /// <summary>Height in points.</summary>
    public float Height { get; set; }
}

/// <summary>Banding and emphasis options for a table.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class TableStyleOptions
{
    /// <summary>Style the first row as a header. None leaves it alone.</summary>
    public TrueFalseNoneEnum HeaderRow { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Style the last row as a totals row. None leaves it alone.</summary>
    public TrueFalseNoneEnum TotalRow { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Emphasise the first column. None leaves it alone.</summary>
    public TrueFalseNoneEnum FirstColumn { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Emphasise the last column. None leaves it alone.</summary>
    public TrueFalseNoneEnum LastColumn { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Band the rows. None leaves it alone.</summary>
    public TrueFalseNoneEnum BandedRows { get; set; } = TrueFalseNoneEnum.None;

    /// <summary>Band the columns. None leaves it alone.</summary>
    public TrueFalseNoneEnum BandedColumns { get; set; } = TrueFalseNoneEnum.None;
}
