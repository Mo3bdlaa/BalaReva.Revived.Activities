using System.Data;
using System.Runtime.InteropServices;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint;

/// <summary>
/// <see cref="IPowerPointService"/> driving PowerPoint through late-bound COM.
/// </summary>
/// <remarks>
/// <para>
/// There is deliberately no interop assembly here. PowerPoint's object model is
/// saturated with <c>Microsoft.Office.Core</c> types — <c>Shape.HasTextFrame</c>,
/// <c>Shape.Type</c>, <c>Table.FirstRow</c> and <c>Font.Bold</c> are all
/// <c>MsoTriState</c> or similar — and even members this code never names need that
/// type resolvable for overload resolution. Microsoft does not publish <c>office.dll</c>
/// on NuGet; the packages that do are third-party repackages, and depending on one
/// would be exactly the supply-chain problem docs/AUDIT.md calls out. Vendoring it is
/// what the published BalaReva packages did.
/// </para>
/// <para>
/// So PowerPoint is driven through its ProgID and <c>dynamic</c>. That costs the
/// compile-time checking a typed interop would give, which is a real loss — but it
/// falls on a file no automated test can reach in any case, since no build agent has
/// PowerPoint installed. The constants below stand in for the enums that would
/// otherwise come from the PIA.
/// </para>
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

/// <summary>
/// The Office and PowerPoint constants this service passes across the COM boundary.
/// </summary>
/// <remarks>
/// These are the values of the PIA enums, named so the call sites read as they would
/// with the typed interop. They are fixed by the Office object model and do not change
/// between versions.
/// </remarks>
internal static class Ppt
{
    // MsoTriState
    internal const int MsoFalse = 0;
    internal const int MsoTrue = -1;
    internal const int MsoCTrue = 1;

    // MsoShapeType
    internal const int MsoLinkedPicture = 11;
    internal const int MsoPicture = 13;

    // PpAlertLevel
    internal const int AlertsNone = 1;
    internal const int AlertsAll = 2;

    // MsoAutomationSecurity
    internal const int SecurityLow = 1;
    internal const int SecurityForceDisable = 3;

    // PpSlideLayout
    internal const int LayoutBlank = 12;

    // PpFixedFormatType
    internal const int FixedFormatPdf = 2;

    // PpPrintOutputType
    internal const int PrintOutputSlides = 1;

    /// <summary>The MsoTriState value for a boolean.</summary>
    internal static int Tri(bool value) => value ? MsoTrue : MsoFalse;
}

/// <summary>A presentation held open through late-bound COM.</summary>
internal sealed class PowerPointPresentation : IPowerPointPresentation
{
    private readonly dynamic _application;
    private readonly dynamic _presentation;
    private bool _closed;

    internal PowerPointPresentation(string filePath, string openPassword, string modifyPassword,
                                    bool displayAlerts, bool macrosEnabled)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("FilePath is required.", nameof(filePath));

        FilePath = filePath;
        var progId = Type.GetTypeFromProgID("PowerPoint.Application")
            ?? throw new InvalidOperationException(
                "PowerPoint is not installed, or its COM automation server is not registered.");

        _application = Activator.CreateInstance(progId)
            ?? throw new InvalidOperationException("PowerPoint's automation server could not be started.");

        try
        {
            _application.DisplayAlerts = displayAlerts ? Ppt.AlertsAll : Ppt.AlertsNone;
            _application.AutomationSecurity = macrosEnabled ? Ppt.SecurityLow : Ppt.SecurityForceDisable;

            // PowerPoint takes passwords appended to the file name rather than as
            // arguments, which is its documented way of opening a protected file.
            var name = filePath;
            if (!string.IsNullOrEmpty(openPassword)) name += $"::{openPassword}::";
            if (!string.IsNullOrEmpty(modifyPassword)) name += $"::{modifyPassword}::";

            _presentation = File.Exists(filePath)
                ? _application.Presentations.Open(name, Ppt.MsoFalse, Ppt.MsoFalse, Ppt.MsoFalse)
                : _application.Presentations.Add(Ppt.MsoFalse);
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
    public int SlideCount() => (int)_presentation.Slides.Count;

    /// <inheritdoc />
    public void NewSlide(int slideIndex)
    {
        var at = Math.Clamp(slideIndex, 1, SlideCount() + 1);
        _presentation.Slides.Add(at, Ppt.LayoutBlank);
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
        _presentation.Slides.Paste(Math.Clamp(slideIndex, 1, SlideCount() + 1));
        Save();
    }

    /// <inheritdoc />
    public string SlideExtractor(int slideIndex)
    {
        var target = Path.Combine(
            Path.GetDirectoryName(Path.GetFullPath(FilePath))!,
            $"{Path.GetFileNameWithoutExtension(FilePath)}_Slide{slideIndex}.pptx");

        Slide(slideIndex).Export(target, "PPTX");
        return target;
    }

    /// <inheritdoc />
    public void HideUnhideSlide(int slideIndex, HideUnhideEnum slideShow)
    {
        Slide(slideIndex).SlideShowTransition.Hidden = Ppt.Tri(slideShow == HideUnhideEnum.Hide);
        Save();
    }

    /// <inheritdoc />
    public void SlideTransitions(int slideIndex, EntryEffectEnum effect, bool onMouseClick, float duration)
    {
        var transition = Slide(slideIndex).SlideShowTransition;
        transition.EntryEffect = (int)effect;
        transition.AdvanceOnClick = Ppt.Tri(onMouseClick);
        if (duration > 0) transition.Duration = duration;
        Save();
    }

    /// <inheritdoc />
    public void Merge(string sourcePpt, int startSlideIndex, int endSlideIndex, int slideAfter)
    {
        if (!File.Exists(sourcePpt))
            throw new FileNotFoundException($"The file '{sourcePpt}' does not exist.", sourcePpt);

        var after = Math.Clamp(slideAfter, 0, SlideCount());
        // Zero for either bound means "all of the source", the only reading that makes
        // an unset range useful rather than an error.
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
            foreach (var shape in TextShapes(index))
            {
                string text = shape.TextFrame.TextRange.Text ?? string.Empty;
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
            foreach (var shape in TextShapes(index))
            {
                var found = shape.TextFrame.TextRange.Find(
                    find, 0, Ppt.Tri(matchCase), Ppt.Tri(wholeWord));
                if (found is null) continue;
                string text = found.Text;
                matches.Add(text);
                table.Rows.Add(index, (string)shape.Name, text);
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
            foreach (var shape in TextShapes(index))
            {
                var range = shape.TextFrame.TextRange;
                var found = range.Replace(find, replace ?? string.Empty, 0,
                    Ppt.Tri(matchCase), Ppt.Tri(wholeWord));

                // Replace returns only the first match, so keep going unless asked not to.
                while (found is not null && !firstOccurrence)
                {
                    int next = (int)found.Start + (int)found.Length;
                    found = range.Replace(find, replace ?? string.Empty, next,
                        Ppt.Tri(matchCase), Ppt.Tri(wholeWord));
                }
            }
        }
        Save();
    }

    /// <inheritdoc />
    public void InsertTextBox(int slideIndex, TextBoxRequest request)
    {
        ArgumentNullException.ThrowIfNull(request);
        var shape = Slide(slideIndex).Shapes.AddTextbox(
            (int)request.Orientation,
            request.Left, request.Top,
            request.Width > 0 ? request.Width : 200f,
            request.Height > 0 ? request.Height : 50f);

        shape.TextFrame.TextRange.Text = request.Text;
        shape.TextFrame.TextRange.ParagraphFormat.Alignment = (int)request.Alignment;
        ApplyStyle(shape.TextFrame.TextRange.Font, request.Style);
        shape.ZOrder((int)request.ZOrder);
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
            foreach (var link in Enumerate(Slide(index).Hyperlinks))
                table.Rows.Add(index, (string)link.TextToDisplay, (string)link.Address,
                               (string)link.SubAddress);
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

        // Width and height of -1 tell PowerPoint to keep the image's own size.
        var shape = Slide(slideIndex).Shapes.AddPicture(
            request.ImagePath, Ppt.MsoFalse, Ppt.MsoCTrue,
            request.Left, request.Top,
            request.Width > 0 ? request.Width : -1f,
            request.Height > 0 ? request.Height : -1f);

        shape.ZOrder((int)request.ZOrder);
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
                (int)format);
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
                (int)format);
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

        Slide(slideIndex).Comments.Add(
            left, top, author ?? string.Empty, initials, commentText ?? string.Empty);
        Save();
    }

    /// <inheritdoc />
    public void CommentsDelete(int slideIndex)
    {
        var comments = Slide(slideIndex).Comments;
        // Backwards: the collection is 1-based and reindexes as each one goes.
        for (var i = (int)comments.Count; i >= 1; i--) comments[i].Delete();
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
        foreach (var comment in Enumerate(Slide(slideIndex).Comments))
        {
            table.Rows.Add(slideIndex, (string)comment.Author, (string)comment.Text, false);
            lines.Add($"{comment.Author}: {comment.Text}");

            if (!includeReplies) continue;
            foreach (var reply in Enumerate(comment.Replies))
            {
                table.Rows.Add(slideIndex, (string)reply.Author, (string)reply.Text, true);
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
        table.FirstCol = Ppt.Tri(request.FirstColumn);
        table.LastCol = Ppt.Tri(request.LastColumn);

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
        for (var r = leaveFirstRow ? 2 : 1; r <= (int)table.Rows.Count; r++)
        {
            for (var c = 1; c <= (int)table.Columns.Count; c++)
                table.Cell(r, c).Shape.TextFrame.TextRange.Text = string.Empty;
        }
        Save();
    }

    /// <inheritdoc />
    public void AppendTable(TableRef reference, string[] values)
    {
        var table = TableShape(reference).Table;
        table.Rows.Add();
        var row = (int)table.Rows.Count;
        var columns = (int)table.Columns.Count;
        for (var c = 0; c < Math.Min(values?.Length ?? 0, columns); c++)
            table.Cell(row, c + 1).Shape.TextFrame.TextRange.Text = values![c] ?? string.Empty;
        Save();
    }

    /// <inheritdoc />
    public void EditTable(TableRef reference, int rowIndex, int columnIndex, string cellValue)
    {
        var table = TableShape(reference).Table;
        Bounds(rowIndex, (int)table.Rows.Count, nameof(rowIndex));
        Bounds(columnIndex, (int)table.Columns.Count, nameof(columnIndex));
        table.Cell(rowIndex, columnIndex).Shape.TextFrame.TextRange.Text = cellValue ?? string.Empty;
        Save();
    }

    /// <inheritdoc />
    public string GetRowItem(TableRef reference, int rowIndex, int columnIndex)
    {
        var table = TableShape(reference).Table;
        Bounds(rowIndex, (int)table.Rows.Count, nameof(rowIndex));
        Bounds(columnIndex, (int)table.Columns.Count, nameof(columnIndex));
        return (string)table.Cell(rowIndex, columnIndex).Shape.TextFrame.TextRange.Text ?? string.Empty;
    }

    /// <inheritdoc />
    public void DeleteRow(TableRef reference, int rowIndex)
    {
        var table = TableShape(reference).Table;
        Bounds(rowIndex, (int)table.Rows.Count, nameof(rowIndex));
        table.Rows[rowIndex].Delete();
        Save();
    }

    /// <inheritdoc />
    public void DeleteColumn(TableRef reference, int columnIndex)
    {
        var table = TableShape(reference).Table;
        Bounds(columnIndex, (int)table.Columns.Count, nameof(columnIndex));
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
        for (var r = 1; r <= (int)table.Rows.Count; r++)
        {
            for (var c = 1; c <= (int)table.Columns.Count; c++)
                ApplyStyle(table.Cell(r, c).Shape.TextFrame.TextRange.Font, style);
        }
        Save();
    }

    /// <inheritdoc />
    public void StyleOption(TableRef reference, TableStyleOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var table = TableShape(reference).Table;
        table.FirstRow = Ppt.Tri(options.HeaderRow);
        table.LastRow = Ppt.Tri(options.TotalRow);
        table.FirstCol = Ppt.Tri(options.FirstColumn);
        table.LastCol = Ppt.Tri(options.LastColumn);
        table.HorizBanding = Ppt.Tri(options.BandedRows);
        table.VertBanding = Ppt.Tri(options.BandedColumns);
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
            string name = shape.Name;
            var data = new DataTable(string.IsNullOrWhiteSpace(name) ? $"Table{index}" : name);

            var columns = (int)table.Columns.Count;
            for (var c = 1; c <= columns; c++) data.Columns.Add($"Column{c}", typeof(string));
            for (var r = 1; r <= (int)table.Rows.Count; r++)
            {
                var values = new object[columns];
                for (var c = 1; c <= columns; c++)
                    values[c - 1] = (string)table.Cell(r, c).Shape.TextFrame.TextRange.Text ?? string.Empty;
                data.Rows.Add(values);
            }
            set.Tables.Add(data);
        }
        return set;
    }

    /// <inheritdoc />
    public string[] GetTableNames(int slideIndex) =>
        [.. TableShapes(slideIndex).Select(shape => (string)shape.Name)];

    // ----------------------------------------------------------------- tools

    /// <inheritdoc />
    public void DataTransformer(int[] slideIndexes, Dictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);
        foreach (var index in Targets(slideIndexes))
        {
            foreach (var shape in TextShapes(index))
            {
                var range = shape.TextFrame.TextRange;
                foreach (var pair in replacements)
                {
                    if (string.IsNullOrEmpty(pair.Key)) continue;
                    var found = range.Replace(pair.Key, pair.Value ?? string.Empty);
                    while (found is not null)
                    {
                        int next = (int)found.Start + (int)found.Length;
                        found = range.Replace(pair.Key, pair.Value ?? string.Empty, next);
                    }
                }
            }
        }
        Save();
    }

    /// <inheritdoc />
    public void ExportTableToExcel(TableRef reference, string excelFile, string sheetName, string startCell)
    {
        // Put the table on the clipboard so the caller still has a route, then say
        // plainly that the rest is not done rather than failing silently.
        TableCopyToClipboard(reference);
        throw new NotSupportedException(
            $"Exporting a table to '{excelFile}' is not implemented. Writing the workbook needs "
            + "the Excel object model, which this package deliberately does not depend on. The "
            + "table has been copied to the clipboard; use BalaReva.Revived.Excel.Activities to "
            + "write it. See docs/REVIVAL.md.");
    }

    /// <inheritdoc />
    public void ImportDataFromExcel(int slideIndex, string excelFile, string sheetName, string cellRange) =>
        throw new NotSupportedException(
            $"Importing from '{excelFile}' is not implemented. Reading the workbook needs the "
            + "Excel object model, which this package deliberately does not depend on. Read the "
            + "range with BalaReva.Revived.Excel.Activities and pass it to AddTable. "
            + "See docs/REVIVAL.md.");

    // ---------------------------------------------------------- presentation

    /// <inheritdoc />
    public void ExportPdf(string filePath, FixedFormatIntentEnum formatType)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("FilePath is required.", nameof(filePath));

        _presentation.ExportAsFixedFormat(filePath, Ppt.FixedFormatPdf, (int)formatType);
    }

    /// <inheritdoc />
    public void SaveAs(string saveAsFile, SaveAsEnum format)
    {
        if (string.IsNullOrWhiteSpace(saveAsFile))
            throw new ArgumentException("SaveAsFile is required.", nameof(saveAsFile));

        _presentation.SaveAs(saveAsFile, (int)format, Ppt.MsoTrue);
    }

    /// <inheritdoc />
    public void Print(int numberOfCopies, PrintColorTypeEnum colorType,
                      bool printComments, bool printHiddenSlides)
    {
        var options = _presentation.PrintOptions;
        options.NumberOfCopies = Math.Max(numberOfCopies, 1);
        options.OutputType = Ppt.PrintOutputSlides;
        options.PrintColorType = (int)colorType;
        options.PrintComments = Ppt.Tri(printComments);
        options.PrintHiddenSlides = Ppt.Tri(printHiddenSlides);
        _presentation.PrintOut();
    }

    /// <inheritdoc />
    public void RemoveDocumentInformation(RemoveDocInfoTypeEnum docInfoType)
    {
        _presentation.RemoveDocumentInformation((int)docInfoType);
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

    private dynamic Slide(int slideIndex)
    {
        Bounds(slideIndex, SlideCount(), nameof(slideIndex));
        return _presentation.Slides[slideIndex];
    }

    /// <summary>The slides an activity means; an empty list means every slide.</summary>
    private IEnumerable<int> Targets(int[] slideIndexes) =>
        slideIndexes is { Length: > 0 } ? slideIndexes : Enumerable.Range(1, SlideCount());

    /// <summary>Walks a COM collection, which is 1-based and not generic.</summary>
    private static IEnumerable<dynamic> Enumerate(dynamic collection)
    {
        var count = (int)collection.Count;
        for (var i = 1; i <= count; i++) yield return collection[i];
    }

    private List<dynamic> Shapes(int slideIndex, Func<dynamic, bool> keep) =>
        [.. Enumerate(Slide(slideIndex).Shapes).Where(keep)];

    private List<dynamic> TextShapes(int slideIndex) =>
        Shapes(slideIndex, shape =>
            (int)shape.HasTextFrame == Ppt.MsoTrue && (int)shape.TextFrame.HasText == Ppt.MsoTrue);

    private List<dynamic> PictureShapes(int slideIndex) =>
        Shapes(slideIndex, shape => (int)shape.Type is Ppt.MsoPicture or Ppt.MsoLinkedPicture);

    private List<dynamic> ChartShapes(int slideIndex) =>
        Shapes(slideIndex, shape => (int)shape.HasChart == Ppt.MsoTrue);

    private List<dynamic> TableShapes(int slideIndex) =>
        Shapes(slideIndex, shape => (int)shape.HasTable == Ppt.MsoTrue);

    /// <summary>
    /// Finds the table an activity meant, preferring the name over the index.
    /// </summary>
    /// <remarks>
    /// A name survives shapes being added and reordered, which an index does not, so it
    /// wins when both are given.
    /// </remarks>
    private dynamic TableShape(TableRef reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        var tables = TableShapes(reference.SlideIndex);

        if (!string.IsNullOrWhiteSpace(reference.TableName))
        {
            foreach (var shape in tables)
            {
                if (string.Equals((string)shape.Name, reference.TableName,
                                  StringComparison.OrdinalIgnoreCase))
                    return shape;
            }
            throw new ArgumentException(
                $"Slide {reference.SlideIndex} has no table named '{reference.TableName}'.",
                nameof(reference));
        }

        Bounds(reference.TableIndex, tables.Count, nameof(reference));
        return tables[reference.TableIndex - 1];
    }

    private static void ApplyStyle(dynamic font, TextStyleRequest? style)
    {
        if (style is null) return;
        if (!string.IsNullOrWhiteSpace(style.FontName)) font.Name = style.FontName;
        if (style.FontSize > 0) font.Size = style.FontSize;

        // None means "leave it alone", so neither branch runs.
        if (style.Bold != TrueFalseNoneEnum.None)
            font.Bold = Ppt.Tri(style.Bold == TrueFalseNoneEnum.True);
        if (style.Italic != TrueFalseNoneEnum.None)
            font.Italic = Ppt.Tri(style.Italic == TrueFalseNoneEnum.True);
    }

    private static string Extension(ImageFileFormatEnum format) => format.ToString().ToLowerInvariant();

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
        try { _application?.Quit(); } catch (COMException) { /* already gone */ }
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
