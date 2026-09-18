using System.Data;
using System.Runtime.InteropServices;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;
using Interop = Microsoft.Office.Interop.PowerPoint;

namespace BalaReva.EasyPowerPoint;

/// <summary>
/// <see cref="IPowerPointService"/> on top of PowerPoint's COM object model.
/// </summary>
/// <remarks>
/// Requires Windows with PowerPoint installed, so this type is never exercised by CI.
///
/// A handful of calls here go through <c>dynamic</c> rather than the typed interop:
/// <c>Presentations.Open</c>, <c>Shapes.AddPicture</c> and <c>SaveAs</c> all take an
/// <c>MsoTriState</c>, which lives in the Office Core PIA. Microsoft does not publish
/// that assembly on NuGet, and the packages that do are third-party repackages. Taking
/// one as a dependency would be exactly the supply-chain problem this repository's own
/// audit calls out, so the few calls that need it are made late-bound instead, passing
/// the underlying values (msoFalse is 0, msoTrue is -1, msoCTrue is 1).
/// </remarks>
public sealed class PowerPointService : IPowerPointService
{
    /// <summary>The shared instance used when a workflow registers no extension.</summary>
    public static PowerPointService Instance { get; } = new();

    /// <inheritdoc />
    public IPowerPointPresentation Open(string filePath, string openPassword, string modifyPassword,
                                        bool displayAlerts, bool macrosEnabled) =>
        new PowerPointPresentation(filePath, openPassword, modifyPassword, displayAlerts, macrosEnabled);
}

/// <summary>A presentation held open through COM.</summary>
internal sealed class PowerPointPresentation : IPowerPointPresentation
{
    private const int MsoFalse = 0;
    private const int MsoTrue = -1;
    private const int MsoCTrue = 1;

    private readonly Interop.Application _application;
    private readonly Interop.Presentation _presentation;
    private bool _closed;

    internal PowerPointPresentation(string filePath, string openPassword, string modifyPassword,
                                    bool displayAlerts, bool macrosEnabled)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("FilePath is required.", nameof(filePath));

        FilePath = filePath;
        _application = new Interop.Application
        {
            DisplayAlerts = displayAlerts
                ? Interop.PpAlertLevel.ppAlertsAll
                : Interop.PpAlertLevel.ppAlertsNone,
        };

        // msoAutomationSecurityLow is 1, msoAutomationSecurityForceDisable is 3.
        dynamic application = _application;
        application.AutomationSecurity = macrosEnabled ? 1 : 3;

        try
        {
            // PowerPoint appends passwords to the file name rather than taking them as
            // arguments, which is its documented way of opening a protected file.
            var name = filePath;
            if (!string.IsNullOrEmpty(openPassword)) name += $"::{openPassword}::";
            if (!string.IsNullOrEmpty(modifyPassword)) name += $"::{modifyPassword}::";

            dynamic presentations = _application.Presentations;
            _presentation = File.Exists(filePath)
                ? (Interop.Presentation)presentations.Open(name, MsoFalse, MsoFalse, MsoFalse)
                : (Interop.Presentation)presentations.Add(MsoFalse);
        }
        catch
        {
            Quit();
            throw;
        }
    }

    /// <inheritdoc />
    public string FilePath { get; }

    // ---------------------------------------------------------------- slides

    /// <inheritdoc />
    public int SlideCount() => _presentation.Slides.Count;

    /// <inheritdoc />
    public void NewSlide(int slideIndex)
    {
        var at = Math.Clamp(slideIndex, 1, _presentation.Slides.Count + 1);
        _presentation.Slides.Add(at, Interop.PpSlideLayout.ppLayoutBlank);
        Save();
    }

    /// <inheritdoc />
    public void DeleteSlide(int slideIndex)
    {
        Slide(slideIndex).Delete();
        Save();
    }

    /// <inheritdoc />
    public void DuplicateSlide(int slideIndex)
    {
        Slide(slideIndex).Duplicate();
        Save();
    }

    /// <inheritdoc />
    public void SlideCopy(int slideIndex) => Slide(slideIndex).Copy();

    /// <inheritdoc />
    public void SlidePaste(int slideIndex)
    {
        _presentation.Slides.Paste(Math.Clamp(slideIndex, 1, _presentation.Slides.Count + 1));
        Save();
    }

    /// <inheritdoc />
    public string SlideExtractor(int slideIndex)
    {
        var slide = Slide(slideIndex);
        var target = Path.Combine(
            Path.GetDirectoryName(Path.GetFullPath(FilePath))!,
            $"{Path.GetFileNameWithoutExtension(FilePath)}_Slide{slideIndex}.pptx");

        slide.Export(target, "PPTX");
        return target;
    }

    /// <inheritdoc />
    public void HideUnhideSlide(int slideIndex, HideUnhideEnum slideShow)
    {
        dynamic transition = Slide(slideIndex).SlideShowTransition;
        transition.Hidden = Tri(slideShow == HideUnhideEnum.Hide);
        Save();
    }

    /// <inheritdoc />
    public void SlideTransitions(int slideIndex, EntryEffectEnum effect, bool onMouseClick, float duration)
    {
        dynamic transition = Slide(slideIndex).SlideShowTransition;
        transition.EntryEffect = (int)effect;
        transition.AdvanceOnClick = Tri(onMouseClick);
        if (duration > 0) transition.Duration = duration;
        Save();
    }

    /// <inheritdoc />
    public void Merge(string sourcePpt, int startSlideIndex, int endSlideIndex, int slideAfter)
    {
        if (!File.Exists(sourcePpt))
            throw new FileNotFoundException($"The file '{sourcePpt}' does not exist.", sourcePpt);

        var after = Math.Clamp(slideAfter, 0, _presentation.Slides.Count);
        // Zero for either bound means "all of the source", which is the only reading
        // that makes an unset range useful.
        if (startSlideIndex <= 0 || endSlideIndex <= 0)
            _presentation.Slides.InsertFromFile(sourcePpt, after);
        else
            _presentation.Slides.InsertFromFile(sourcePpt, after, startSlideIndex, endSlideIndex);
        Save();
    }

    // -------------------------------------------------------- text and shapes

    /// <inheritdoc />
    public (string[] Array, string Text) ReadText(int[] slideIndexes, bool addSlideIndex, bool omitEmptyLine)
    {
        var lines = new List<string>();
        foreach (var index in Targets(slideIndexes))
        {
            foreach (Interop.Shape shape in Slide(index).Shapes)
            {
                if (!HasText(shape)) continue;
                var text = shape.TextFrame.TextRange.Text ?? string.Empty;
                foreach (var line in text.Replace("\r\n", "\r").Split('\r'))
                {
                    if (omitEmptyLine && string.IsNullOrWhiteSpace(line)) continue;
                    lines.Add(addSlideIndex ? $"{index}\t{line}" : line);
                }
            }
        }
        return ([.. lines], string.Join(Environment.NewLine, lines));
    }

    /// <inheritdoc />
    public (string[] Array, DataTable Table) FindText(int[] slideIndexes, string find,
                                                      bool matchCase, bool wholeWord)
    {
        if (string.IsNullOrEmpty(find))
            throw new ArgumentException("FindString is required.", nameof(find));

        var matches = new List<string>();
        var table = new DataTable("FindText");
        table.Columns.Add("SlideIndex", typeof(int));
        table.Columns.Add("ShapeName", typeof(string));
        table.Columns.Add("Text", typeof(string));

        foreach (var index in Targets(slideIndexes))
        {
            foreach (Interop.Shape shape in Slide(index).Shapes)
            {
                if (!HasText(shape)) continue;
                var found = shape.TextFrame.TextRange.Find(find, 0, Tri(matchCase), Tri(wholeWord));
                if (found is null) continue;
                matches.Add(found.Text);
                table.Rows.Add(index, shape.Name, found.Text);
            }
        }
        return ([.. matches], table);
    }

    /// <inheritdoc />
    public void FindReplace(int[] slideIndexes, string find, string replace,
                            bool matchCase, bool wholeWord, bool firstOccurrence)
    {
        if (string.IsNullOrEmpty(find))
            throw new ArgumentException("FindText is required.", nameof(find));

        foreach (var index in Targets(slideIndexes))
        {
            foreach (Interop.Shape shape in Slide(index).Shapes)
            {
                if (!HasText(shape)) continue;
                var range = shape.TextFrame.TextRange;
                var found = range.Replace(find, replace ?? string.Empty, 0, Tri(matchCase), Tri(wholeWord));
                // Replace returns only the first match, so keep going unless asked not to.
                while (found is not null && !firstOccurrence)
                {
                    found = range.Replace(find, replace ?? string.Empty,
                        found.Start + found.Length, Tri(matchCase), Tri(wholeWord));
                }
            }
        }
        Save();
    }

    /// <inheritdoc />
    public void InsertTextBox(int slideIndex, TextBoxRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        dynamic shapes = Slide(slideIndex).Shapes;
        var shape = (Interop.Shape)shapes.AddTextbox(
            (int)request.Orientation,
            request.Left, request.Top,
            request.Width > 0 ? request.Width : 200f,
            request.Height > 0 ? request.Height : 50f);

        shape.TextFrame.TextRange.Text = request.Text;
        shape.TextFrame.TextRange.ParagraphFormat.Alignment =
            (Interop.PpParagraphAlignment)(int)request.Alignment;
        ApplyStyle(shape.TextFrame.TextRange.Font, request.Style);
        ZOrder(shape, request.ZOrder);
        Save();
    }

    /// <inheritdoc />
    public void TextShapeEdit(int slideIndex, int textIndex, TextStyleRequest style)
    {
        var shapes = TextShapes(slideIndex);
        Bounds(textIndex, shapes.Count, nameof(textIndex));
        ApplyStyle(shapes[textIndex - 1].TextFrame.TextRange.Font, style);
        Save();
    }

    /// <inheritdoc />
    public int TextShapeCount(int slideIndex) => TextShapes(slideIndex).Count;

    /// <inheritdoc />
    public DataTable ExtractHyperLinks(int[] slideIndexes)
    {
        var table = new DataTable("HyperLinks");
        table.Columns.Add("SlideIndex", typeof(int));
        table.Columns.Add("TextToDisplay", typeof(string));
        table.Columns.Add("Address", typeof(string));
        table.Columns.Add("SubAddress", typeof(string));

        foreach (var index in Targets(slideIndexes))
        {
            foreach (Interop.Hyperlink link in Slide(index).Hyperlinks)
            {
                table.Rows.Add(index, link.TextToDisplay, link.Address, link.SubAddress);
            }
        }
        return table;
    }

    // ---------------------------------------------------------------- images

    /// <inheritdoc />
    public void InsertPicture(int slideIndex, PictureRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        if (!File.Exists(request.ImagePath))
        {
            throw new FileNotFoundException(
                $"The image '{request.ImagePath}' does not exist.", request.ImagePath);
        }

        dynamic shapes = Slide(slideIndex).Shapes;
        // Width and height of -1 tell PowerPoint to keep the image's own size.
        var shape = (Interop.Shape)shapes.AddPicture(
            request.ImagePath, MsoFalse, MsoCTrue,
            request.Left, request.Top,
            request.Width > 0 ? request.Width : -1f,
            request.Height > 0 ? request.Height : -1f);

        ZOrder(shape, request.ZOrder);
        Save();
    }

    /// <inheritdoc />
    public void DeleteImage(int slideIndex, int imageIndex)
    {
        var pictures = PictureShapes(slideIndex);
        Bounds(imageIndex, pictures.Count, nameof(imageIndex));
        pictures[imageIndex - 1].Delete();
        Save();
    }

    /// <inheritdoc />
    public int ImageShapeCount(int slideIndex) => PictureShapes(slideIndex).Count;

    /// <inheritdoc />
    public void ImageExtractor(int slideIndex, string imageDirectory, ImageFileFormatEnum format)
    {
        if (string.IsNullOrWhiteSpace(imageDirectory))
            throw new ArgumentException("ImageDirectory is required.", nameof(imageDirectory));
        Directory.CreateDirectory(imageDirectory);

        var pictures = PictureShapes(slideIndex);
        for (var i = 0; i < pictures.Count; i++)
        {
            pictures[i].Export(
                Path.Combine(imageDirectory, $"Slide{slideIndex}_Image{i + 1}.{Extension(format)}"),
                (Interop.PpShapeFormat)(int)format);
        }
    }

    /// <inheritdoc />
    public void PasteClipboard(int slideIndex, double left, double top, double width, double height)
    {
        var pasted = Slide(slideIndex).Shapes.Paste();
        if (left > 0) pasted.Left = (float)left;
        if (top > 0) pasted.Top = (float)top;
        if (width > 0) pasted.Width = (float)width;
        if (height > 0) pasted.Height = (float)height;
        Save();
    }

    // ---------------------------------------------------------------- charts

    /// <inheritdoc />
    public int ChartShapeCount(int slideIndex) => ChartShapes(slideIndex).Count;

    /// <inheritdoc />
    public void ChartDelete(int slideIndex, int chartIndex)
    {
        var charts = ChartShapes(slideIndex);
        Bounds(chartIndex, charts.Count, nameof(chartIndex));
        charts[chartIndex - 1].Delete();
        Save();
    }

    /// <inheritdoc />
    public void ChartCopyToClipboard(int slideIndex, int chartIndex)
    {
        var charts = ChartShapes(slideIndex);
        Bounds(chartIndex, charts.Count, nameof(chartIndex));
        charts[chartIndex - 1].Copy();
    }

    /// <inheritdoc />
    public void ChartFormat(int slideIndex, int chartIndex, double left, double top,
                            double width, double height)
    {
        var charts = ChartShapes(slideIndex);
        Bounds(chartIndex, charts.Count, nameof(chartIndex));
        var shape = charts[chartIndex - 1];
        if (left > 0) shape.Left = (float)left;
        if (top > 0) shape.Top = (float)top;
        if (width > 0) shape.Width = (float)width;
        if (height > 0) shape.Height = (float)height;
        Save();
    }

    /// <inheritdoc />
    public void ChartImageExtract(int slideIndex, string imageFolder, ImageFileFormatEnum format)
    {
        if (string.IsNullOrWhiteSpace(imageFolder))
            throw new ArgumentException("ImageFolder is required.", nameof(imageFolder));
        Directory.CreateDirectory(imageFolder);

        var charts = ChartShapes(slideIndex);
        for (var i = 0; i < charts.Count; i++)
        {
            charts[i].Export(
                Path.Combine(imageFolder, $"Slide{slideIndex}_Chart{i + 1}.{Extension(format)}"),
                (Interop.PpShapeFormat)(int)format);
        }
    }

    /// <inheritdoc />
    public void RefreshData(int slideIndex)
    {
        foreach (var shape in ChartShapes(slideIndex)) shape.Chart.Refresh();
        Save();
    }

    /// <inheritdoc />
    public void UpdateLinks()
    {
        _presentation.UpdateLinks();
        Save();
    }

    // -------------------------------------------------------------- comments

    /// <inheritdoc />
    public void CommentsAdd(int slideIndex, string author, string commentText, float left, float top)
    {
        var initials = string.IsNullOrWhiteSpace(author)
            ? "?"
            : string.Concat(author.Split(' ', StringSplitOptions.RemoveEmptyEntries)
                                  .Select(part => part[0]));

        Slide(slideIndex).Comments.Add(left, top, author ?? string.Empty, initials, commentText ?? string.Empty);
        Save();
    }

    /// <inheritdoc />
    public void CommentsDelete(int slideIndex)
    {
        var comments = Slide(slideIndex).Comments;
        // Backwards: the collection is 1-based and reindexes as each one goes.
        for (var i = comments.Count; i >= 1; i--) comments[i].Delete();
        Save();
    }

    /// <inheritdoc />
    public (string Text, DataTable Table) CommentsRead(int slideIndex, bool includeReplies)
    {
        var table = new DataTable("Comments");
        table.Columns.Add("SlideIndex", typeof(int));
        table.Columns.Add("Author", typeof(string));
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("IsReply", typeof(bool));

        var lines = new List<string>();
        foreach (Interop.Comment comment in Slide(slideIndex).Comments)
        {
            table.Rows.Add(slideIndex, comment.Author, comment.Text, false);
            lines.Add($"{comment.Author}: {comment.Text}");

            if (!includeReplies) continue;
            foreach (Interop.Comment reply in comment.Replies)
            {
                table.Rows.Add(slideIndex, reply.Author, reply.Text, true);
                lines.Add($"    {reply.Author}: {reply.Text}");
            }
        }
        return (string.Join(Environment.NewLine, lines), table);
    }

    // ---------------------------------------------------------------- tables

    /// <inheritdoc />
    public void AddTable(int slideIndex, AddTableRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var input = request.InputTable;
        var rows = input.Rows.Count + (request.AddHeader ? 1 : 0);
        var columns = input.Columns.Count;
        if (rows == 0 || columns == 0)
            throw new ArgumentException("InputTable has no rows or no columns.", nameof(request));

        var shape = Slide(slideIndex).Shapes.AddTable(
            rows, columns, request.Left, request.Top,
            request.Width > 0 ? request.Width : -1f,
            request.Height > 0 ? request.Height : -1f);

        if (!string.IsNullOrWhiteSpace(request.TableName)) shape.Name = request.TableName;
        var table = shape.Table;
        dynamic flags = table;
        flags.FirstCol = Tri(request.FirstColumn);
        flags.LastCol = Tri(request.LastColumn);

        var offset = 1;
        if (request.AddHeader)
        {
            for (var c = 0; c < columns; c++)
                table.Cell(1, c + 1).Shape.TextFrame.TextRange.Text = input.Columns[c].ColumnName;
            offset = 2;
        }
        for (var r = 0; r < input.Rows.Count; r++)
        {
            for (var c = 0; c < columns; c++)
            {
                table.Cell(r + offset, c + 1).Shape.TextFrame.TextRange.Text =
                    input.Rows[r][c]?.ToString() ?? string.Empty;
            }
        }
        Save();
    }

    /// <inheritdoc />
    public void DeleteTable(TableRef reference)
    {
        TableShape(reference).Delete();
        Save();
    }

    /// <inheritdoc />
    public void ClearTable(TableRef reference, bool leaveFirstRow)
    {
        var table = TableShape(reference).Table;
        for (var r = leaveFirstRow ? 2 : 1; r <= table.Rows.Count; r++)
        {
            for (var c = 1; c <= table.Columns.Count; c++)
                table.Cell(r, c).Shape.TextFrame.TextRange.Text = string.Empty;
        }
        Save();
    }

    /// <inheritdoc />
    public void AppendTable(TableRef reference, string[] values)
    {
        var table = TableShape(reference).Table;
        table.Rows.Add();
        var row = table.Rows.Count;
        for (var c = 0; c < Math.Min(values?.Length ?? 0, table.Columns.Count); c++)
            table.Cell(row, c + 1).Shape.TextFrame.TextRange.Text = values![c] ?? string.Empty;
        Save();
    }

    /// <inheritdoc />
    public void EditTable(TableRef reference, int rowIndex, int columnIndex, string cellValue)
    {
        var table = TableShape(reference).Table;
        Bounds(rowIndex, table.Rows.Count, nameof(rowIndex));
        Bounds(columnIndex, table.Columns.Count, nameof(columnIndex));
        table.Cell(rowIndex, columnIndex).Shape.TextFrame.TextRange.Text = cellValue ?? string.Empty;
        Save();
    }

    /// <inheritdoc />
    public string GetRowItem(TableRef reference, int rowIndex, int columnIndex)
    {
        var table = TableShape(reference).Table;
        Bounds(rowIndex, table.Rows.Count, nameof(rowIndex));
        Bounds(columnIndex, table.Columns.Count, nameof(columnIndex));
        return table.Cell(rowIndex, columnIndex).Shape.TextFrame.TextRange.Text ?? string.Empty;
    }

    /// <inheritdoc />
    public void DeleteRow(TableRef reference, int rowIndex)
    {
        var table = TableShape(reference).Table;
        Bounds(rowIndex, table.Rows.Count, nameof(rowIndex));
        table.Rows[rowIndex].Delete();
        Save();
    }

    /// <inheritdoc />
    public void DeleteColumn(TableRef reference, int columnIndex)
    {
        var table = TableShape(reference).Table;
        Bounds(columnIndex, table.Columns.Count, nameof(columnIndex));
        table.Columns[columnIndex].Delete();
        Save();
    }

    /// <inheritdoc />
    public void ResizeTable(TableRef reference, double left, double top, double width, double height)
    {
        var shape = TableShape(reference);
        if (left > 0) shape.Left = (float)left;
        if (top > 0) shape.Top = (float)top;
        if (width > 0) shape.Width = (float)width;
        if (height > 0) shape.Height = (float)height;
        Save();
    }

    /// <inheritdoc />
    public void FontOption(TableRef reference, TextStyleRequest style)
    {
        var table = TableShape(reference).Table;
        for (var r = 1; r <= table.Rows.Count; r++)
        {
            for (var c = 1; c <= table.Columns.Count; c++)
                ApplyStyle(table.Cell(r, c).Shape.TextFrame.TextRange.Font, style);
        }
        Save();
    }

    /// <inheritdoc />
    public void StyleOption(TableRef reference, TableStyleOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        dynamic table = TableShape(reference).Table;
        table.FirstRow = Tri(options.HeaderRow);
        table.LastRow = Tri(options.TotalRow);
        table.FirstCol = Tri(options.FirstColumn);
        table.LastCol = Tri(options.LastColumn);
        table.HorizBanding = Tri(options.BandedRows);
        table.VertBanding = Tri(options.BandedColumns);
        Save();
    }

    /// <inheritdoc />
    public void TableCopyToClipboard(TableRef reference) => TableShape(reference).Copy();

    /// <inheritdoc />
    public DataSet ExtractTables(int slideIndex)
    {
        var set = new DataSet("SlideTables");
        var index = 0;
        foreach (var shape in TableShapes(slideIndex))
        {
            index++;
            var table = shape.Table;
            var data = new DataTable(string.IsNullOrWhiteSpace(shape.Name) ? $"Table{index}" : shape.Name);
            for (var c = 1; c <= table.Columns.Count; c++) data.Columns.Add($"Column{c}", typeof(string));
            for (var r = 1; r <= table.Rows.Count; r++)
            {
                var values = new object[table.Columns.Count];
                for (var c = 1; c <= table.Columns.Count; c++)
                    values[c - 1] = table.Cell(r, c).Shape.TextFrame.TextRange.Text ?? string.Empty;
                data.Rows.Add(values);
            }
            set.Tables.Add(data);
        }
        return set;
    }

    /// <inheritdoc />
    public string[] GetTableNames(int slideIndex) =>
        [.. TableShapes(slideIndex).Select(shape => shape.Name)];

    // ----------------------------------------------------------------- tools

    /// <inheritdoc />
    public void DataTransformer(int[] slideIndexes, Dictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);
        foreach (var index in Targets(slideIndexes))
        {
            foreach (Interop.Shape shape in Slide(index).Shapes)
            {
                if (!HasText(shape)) continue;
                var range = shape.TextFrame.TextRange;
                foreach (var pair in replacements)
                {
                    if (string.IsNullOrEmpty(pair.Key)) continue;
                    var found = range.Replace(pair.Key, pair.Value ?? string.Empty);
                    while (found is not null)
                    {
                        found = range.Replace(pair.Key, pair.Value ?? string.Empty,
                            found.Start + found.Length);
                    }
                }
            }
        }
        Save();
    }

    /// <inheritdoc />
    public void ExportTableToExcel(TableRef reference, string excelFile, string sheetName, string startCell)
    {
        // Writes the table out through PowerPoint's clipboard rather than driving Excel,
        // which would mean a second Office dependency for one activity.
        TableShape(reference).Copy();
        throw new NotSupportedException(
            $"Exporting a table to '{excelFile}' is not implemented. It would need the Excel "
            + "object model, which this package deliberately does not depend on. "
            + "Use the table on the clipboard, or BalaReva.Revived.Excel.Activities. "
            + "See docs/REVIVAL.md.");
    }

    /// <inheritdoc />
    public void ImportDataFromExcel(int slideIndex, string excelFile, string sheetName, string cellRange) =>
        throw new NotSupportedException(
            $"Importing from '{excelFile}' is not implemented. It would need the Excel object "
            + "model, which this package deliberately does not depend on. "
            + "Read the range with BalaReva.Revived.Excel.Activities and pass it to AddTable. "
            + "See docs/REVIVAL.md.");

    // ---------------------------------------------------------- presentation

    /// <inheritdoc />
    public void ExportPdf(string filePath, FixedFormatIntentEnum formatType)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("FilePath is required.", nameof(filePath));

        _presentation.ExportAsFixedFormat(
            filePath,
            Interop.PpFixedFormatType.ppFixedFormatTypePDF,
            (Interop.PpFixedFormatIntent)(int)formatType);
    }

    /// <inheritdoc />
    public void SaveAs(string saveAsFile, SaveAsEnum format)
    {
        if (string.IsNullOrWhiteSpace(saveAsFile))
            throw new ArgumentException("SaveAsFile is required.", nameof(saveAsFile));

        dynamic presentation = _presentation;
        presentation.SaveAs(saveAsFile, (int)format, MsoTrue);
    }

    /// <inheritdoc />
    public void Print(int numberOfCopies, PrintColorTypeEnum colorType,
                      bool printComments, bool printHiddenSlides)
    {
        var options = _presentation.PrintOptions;
        options.NumberOfCopies = Math.Max(numberOfCopies, 1);
        options.OutputType = Interop.PpPrintOutputType.ppPrintOutputSlides;
        options.PrintColorType = (Interop.PpPrintColorType)(int)colorType;
        dynamic flags = options;
        flags.PrintComments = Tri(printComments);
        flags.PrintHiddenSlides = Tri(printHiddenSlides);
        _presentation.PrintOut();
    }

    /// <inheritdoc />
    public void RemoveDocumentInformation(RemoveDocInfoTypeEnum docInfoType)
    {
        _presentation.RemoveDocumentInformation((Interop.PpRemoveDocInfoType)(int)docInfoType);
        Save();
    }

    /// <inheritdoc />
    public object? ExecuteMacro(string macroName, object[]? arguments)
    {
        if (string.IsNullOrWhiteSpace(macroName))
            throw new ArgumentException("MacroName is required.", nameof(macroName));

        return _application.Run(macroName, arguments ?? []);
    }

    // -------------------------------------------------------------- plumbing

    private Interop.Slide Slide(int slideIndex)
    {
        Bounds(slideIndex, _presentation.Slides.Count, nameof(slideIndex));
        return _presentation.Slides[slideIndex];
    }

    /// <summary>The slides an activity means; an empty list means every slide.</summary>
    private IEnumerable<int> Targets(int[] slideIndexes) =>
        slideIndexes is { Length: > 0 }
            ? slideIndexes
            : Enumerable.Range(1, _presentation.Slides.Count);

    private static bool HasText(Interop.Shape shape)
    {
        dynamic s = shape;
        return (int)s.HasTextFrame == MsoTrue && (int)s.TextFrame.HasText == MsoTrue;
    }

    /// <summary>True when the shape carries a chart.</summary>
    private static bool HasChart(Interop.Shape shape)
    {
        dynamic s = shape;
        return (int)s.HasChart == MsoTrue;
    }

    /// <summary>True when the shape carries a table.</summary>
    private static bool HasTable(Interop.Shape shape)
    {
        dynamic s = shape;
        return (int)s.HasTable == MsoTrue;
    }

    /// <summary>True when the shape is a picture, linked or embedded.</summary>
    private static bool IsPicture(Interop.Shape shape)
    {
        dynamic s = shape;
        // msoPicture is 13, msoLinkedPicture is 11.
        var type = (int)s.Type;
        return type is 13 or 11;
    }

    private List<Interop.Shape> TextShapes(int slideIndex) =>
        [.. Slide(slideIndex).Shapes.Cast<Interop.Shape>().Where(HasText)];

    private List<Interop.Shape> PictureShapes(int slideIndex) =>
        [.. Slide(slideIndex).Shapes.Cast<Interop.Shape>().Where(IsPicture)];

    private List<Interop.Shape> ChartShapes(int slideIndex) =>
        [.. Slide(slideIndex).Shapes.Cast<Interop.Shape>().Where(HasChart)];

    private List<Interop.Shape> TableShapes(int slideIndex) =>
        [.. Slide(slideIndex).Shapes.Cast<Interop.Shape>().Where(HasTable)];

    /// <summary>
    /// Finds the table an activity meant, preferring the name over the index.
    /// </summary>
    /// <remarks>
    /// A name survives slides being reordered and shapes being added, which an index
    /// does not, so it wins when both are given.
    /// </remarks>
    private Interop.Shape TableShape(TableRef reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        var tables = TableShapes(reference.SlideIndex);

        if (!string.IsNullOrWhiteSpace(reference.TableName))
        {
            return tables.FirstOrDefault(
                       s => string.Equals(s.Name, reference.TableName, StringComparison.OrdinalIgnoreCase))
                   ?? throw new ArgumentException(
                       $"Slide {reference.SlideIndex} has no table named '{reference.TableName}'.",
                       nameof(reference));
        }

        Bounds(reference.TableIndex, tables.Count, nameof(reference));
        return tables[reference.TableIndex - 1];
    }

    private static void ApplyStyle(Interop.Font font, TextStyleRequest? style)
    {
        if (style is null) return;
        if (!string.IsNullOrWhiteSpace(style.FontName)) font.Name = style.FontName;
        if (style.FontSize > 0) font.Size = style.FontSize;

        // None means "leave it alone", so neither branch runs.
        dynamic f = font;
        if (style.Bold != TrueFalseNoneEnum.None) f.Bold = Tri(style.Bold == TrueFalseNoneEnum.True);
        if (style.Italic != TrueFalseNoneEnum.None) f.Italic = Tri(style.Italic == TrueFalseNoneEnum.True);
    }

    /// <summary>Moves a shape in the slide's stacking order.</summary>
    private static void ZOrder(Interop.Shape shape, ZOrderCmdEnum command)
    {
        dynamic s = shape;
        s.ZOrder((int)command);
    }

    /// <summary>The MsoTriState value for a boolean, as the int the COM call expects.</summary>
    private static int Tri(bool value) => value ? MsoTrue : MsoFalse;

    private static string Extension(ImageFileFormatEnum format) =>
        format.ToString().ToLowerInvariant();

    private static void Bounds(int index, int count, string name)
    {
        if (index < 1 || index > count)
            throw new ArgumentOutOfRangeException(name, index, $"Must be between 1 and {count}.");
    }

    private void Save()
    {
        if (File.Exists(FilePath)) _presentation.Save();
        else _presentation.SaveAs(FilePath);
    }

    private void Quit()
    {
        try { _application.Quit(); } catch (COMException) { /* already gone */ }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_closed) return;
        _closed = true;
        try { _presentation?.Close(); } catch (COMException) { /* nothing to close */ }
        Quit();
    }
}
