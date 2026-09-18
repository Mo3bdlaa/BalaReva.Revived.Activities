using System.Data;
using System.Runtime.InteropServices;
using BalaReva.Word.Utilities;
using Interop = Microsoft.Office.Interop.Word;

namespace BalaReva.Word;

/// <summary>
/// <see cref="IWordService"/> on top of Word's COM object model.
/// </summary>
/// <remarks>
/// Requires Windows with Word installed, so this type is never exercised by CI. Every
/// method is a mechanical translation; the judgement lives in the activities, where a
/// stand-in can cover it.
/// </remarks>
public sealed class WordService : IWordService
{
    /// <summary>The shared instance used when a workflow registers no extension.</summary>
    public static WordService Instance { get; } = new();

    /// <inheritdoc />
    public IWordDocument Open(string filePath, string openPassword, string modifyPassword) =>
        new WordDocument(filePath, openPassword, modifyPassword);

    /// <inheritdoc />
    public void MergeDocuments(
        string wordFile, string openPassword, string modifyPassword, string[] sourceFiles)
    {
        ArgumentNullException.ThrowIfNull(sourceFiles);
        using var document = new WordDocument(wordFile, openPassword, modifyPassword);

        foreach (var source in sourceFiles)
        {
            if (string.IsNullOrWhiteSpace(source)) continue;
            if (!File.Exists(source))
                throw new FileNotFoundException($"The file '{source}' does not exist.", source);

            // Each source starts on its own page; without the break they run together.
            var end = document.Native.Content.End - 1;
            var range = document.Native.Range(end, end);
            range.InsertBreak(Interop.WdBreakType.wdPageBreak);
            range.InsertFile(source);
        }

        document.Save();
    }

    /// <inheritdoc />
    public void WordToPdf(string wordFile, string openPassword, string modifyPassword,
                          string pdfFile, int startPage, int endPage)
    {
        if (string.IsNullOrWhiteSpace(pdfFile))
            throw new ArgumentException("PDFFile is required.", nameof(pdfFile));

        using var document = new WordDocument(wordFile, openPassword, modifyPassword);

        // Zero on either bound means "the whole document"; Word wants a range kind
        // rather than a sentinel page number.
        var wholeDocument = startPage <= 0 || endPage <= 0;
        document.Native.ExportAsFixedFormat(
            pdfFile,
            Interop.WdExportFormat.wdExportFormatPDF,
            Range: wholeDocument
                ? Interop.WdExportRange.wdExportAllDocument
                : Interop.WdExportRange.wdExportFromTo,
            From: wholeDocument ? 1 : startPage,
            To: wholeDocument ? 1 : endPage);
    }
}

/// <summary>A Word document held open through COM.</summary>
internal sealed class WordDocument : IWordDocument
{
    private readonly Interop.Application _application;
    private bool _closed;

    internal WordDocument(string filePath, string openPassword, string modifyPassword)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("FilePath is required.", nameof(filePath));

        Target = new WordObject
        {
            FilePath = filePath,
            Password = openPassword ?? string.Empty,
            ModiPassword = modifyPassword ?? string.Empty,
        };

        _application = new Interop.Application { Visible = false, DisplayAlerts = Interop.WdAlertLevel.wdAlertsNone };
        try
        {
            Native = File.Exists(filePath)
                ? _application.Documents.Open(filePath,
                    PasswordDocument: Target.Password,
                    WritePasswordDocument: Target.ModiPassword)
                : _application.Documents.Add();
        }
        catch
        {
            Quit();
            throw;
        }
    }

    /// <summary>The underlying Word document.</summary>
    internal Interop.Document Native { get; private set; }

    /// <inheritdoc />
    public WordObject Target { get; }

    /// <summary>Saves the document back to its own path.</summary>
    internal void Save() => Native.Save();

    // ----------------------------------------------------------- documents

    /// <inheritdoc />
    public void ChangePassword(string newOpenPassword, string newModifyPassword)
    {
        Native.Password = newOpenPassword ?? string.Empty;
        Native.WritePassword = newModifyPassword ?? string.Empty;
        Native.Save();
    }

    /// <inheritdoc />
    public void CreateDocument()
    {
        var created = _application.Documents.Add();
        created.SaveAs2(Target.FilePath);
        Native = created;
    }

    /// <inheritdoc />
    public void SaveAs(string newFileName, EnumSaveAs format)
    {
        if (string.IsNullOrWhiteSpace(newFileName))
            throw new ArgumentException("NewFileName is required.", nameof(newFileName));

        Native.SaveAs2(newFileName, (Interop.WdSaveFormat)(int)format);
    }

    /// <inheritdoc />
    public void Print(string printerName, int copies, EnumOrientation orientation)
    {
        if (!string.IsNullOrWhiteSpace(printerName)) _application.ActivePrinter = printerName;
        if (orientation != EnumOrientation.Select)
        {
            Native.PageSetup.Orientation = (Interop.WdOrientation)(int)orientation;
        }

        Native.PrintOut(Copies: Math.Max(copies, 1));
    }

    // ------------------------------------------------------------- content

    /// <inheritdoc />
    public void Select(EnumGoTo target, string bookmark, int pageOrLine, int rows, int columns)
    {
        GoTo(target, bookmark, pageOrLine);
        if (rows > 0 && columns > 0)
        {
            // A row and column count only makes sense when the caller means a table
            // cell; extending the selection is what the published activity's arguments
            // imply, and doing nothing would be a silent no-op.
            _application.Selection.MoveDown(Interop.WdUnits.wdLine, rows, Interop.WdMovementType.wdExtend);
            _application.Selection.MoveRight(Interop.WdUnits.wdCharacter, columns, Interop.WdMovementType.wdExtend);
        }
    }

    /// <inheritdoc />
    public void Paste(EnumGoTo target, string bookmark, int pageOrLine)
    {
        GoTo(target, bookmark, pageOrLine);
        _application.Selection.Paste();
    }

    /// <inheritdoc />
    public void AddPageBreak() =>
        _application.Selection.InsertBreak(Interop.WdBreakType.wdPageBreak);

    /// <inheritdoc />
    public void RemoveDisplayLineNumber()
    {
        foreach (Interop.Section section in Native.Sections)
            section.PageSetup.LineNumbering.Active = 0;
    }

    /// <inheritdoc />
    public void FindReplace(string findText, string replaceText, EnumFindReplace findOption,
                            EnumReplaceOption replaceOption, bool matchCase)
    {
        if (string.IsNullOrEmpty(findText))
            throw new ArgumentException("FindText is required.", nameof(findText));

        var find = Native.Content.Find;
        find.ClearFormatting();
        find.Replacement.ClearFormatting();
        find.Text = findText;
        find.Replacement.Text = replaceText ?? string.Empty;
        find.MatchCase = matchCase;
        find.MatchWholeWord = findOption == EnumFindReplace.Whole;
        find.Forward = true;
        find.Wrap = Interop.WdFindWrap.wdFindContinue;

        find.Execute(Replace: replaceOption == EnumReplaceOption.ReplaceAll
            ? Interop.WdReplace.wdReplaceAll
            : Interop.WdReplace.wdReplaceOne);
    }

    /// <inheritdoc />
    public object? ExecuteMacro(string macroName, object[]? parameters)
    {
        if (string.IsNullOrWhiteSpace(macroName))
            throw new ArgumentException("MacroName is required.", nameof(macroName));

        var arguments = new object[(parameters?.Length ?? 0) + 1];
        arguments[0] = macroName;
        parameters?.CopyTo(arguments, 1);

        return _application.GetType().InvokeMember(
            "Run", System.Reflection.BindingFlags.InvokeMethod, null, _application, arguments);
    }

    /// <inheritdoc />
    public WordStatisticsResult Statistics() => new()
    {
        Characters = Native.ComputeStatistics(Interop.WdStatistic.wdStatisticCharacters),
        CharactersWithSpaces = Native.ComputeStatistics(Interop.WdStatistic.wdStatisticCharactersWithSpaces),
        Lines = Native.ComputeStatistics(Interop.WdStatistic.wdStatisticLines),
        Pages = Native.ComputeStatistics(Interop.WdStatistic.wdStatisticPages),
        Paragraphs = Native.ComputeStatistics(Interop.WdStatistic.wdStatisticParagraphs),
        WordCount = Native.ComputeStatistics(Interop.WdStatistic.wdStatisticWords),
    };

    /// <inheritdoc />
    public void ExtractImages(string imageFolder, EnumFileExtension extension)
    {
        if (string.IsNullOrWhiteSpace(imageFolder))
            throw new ArgumentException("ImageFolder is required.", nameof(imageFolder));
        Directory.CreateDirectory(imageFolder);

        var index = 0;
        foreach (Interop.InlineShape shape in Native.InlineShapes)
        {
            index++;
            shape.Range.CopyAsPicture();
            SaveClipboardImage(Path.Combine(
                imageFolder, $"Image{index}.{extension.ToString().ToLowerInvariant()}"), extension);
        }
    }

    // --------------------------------------------------- headers and footers

    /// <inheritdoc />
    public string[] ReadHeaderFooter(EnumHeadersFooters readType)
    {
        var results = new List<string>();
        foreach (Interop.Section section in Native.Sections)
        {
            var parts = readType == EnumHeadersFooters.Header ? section.Headers : section.Footers;
            foreach (Interop.HeaderFooter part in parts)
            {
                var text = part.Range.Text;
                if (!string.IsNullOrWhiteSpace(text)) results.Add(text.TrimEnd('\r'));
            }
        }
        return [.. results];
    }

    /// <inheritdoc />
    public void InsertHeaderFooter(EnumHeadersFooters insertType, HeaderFooterText text)
    {
        ArgumentNullException.ThrowIfNull(text);
        foreach (Interop.Section section in Native.Sections)
        {
            Write(section, insertType, Interop.WdHeaderFooterIndex.wdHeaderFooterPrimary, text.OddPageText, text);
            Write(section, insertType, Interop.WdHeaderFooterIndex.wdHeaderFooterEvenPages, text.EvenPageText, text);
            Write(section, insertType, Interop.WdHeaderFooterIndex.wdHeaderFooterFirstPage, text.FirstPageText, text);
        }
    }

    private static void Write(Interop.Section section, EnumHeadersFooters kind,
                              Interop.WdHeaderFooterIndex index, string value, HeaderFooterText format)
    {
        if (string.IsNullOrEmpty(value)) return;

        var parts = kind == EnumHeadersFooters.Header ? section.Headers : section.Footers;
        var range = parts[index].Range;
        range.Text = value;

        if (!string.IsNullOrWhiteSpace(format.FontName)) range.Font.Name = format.FontName;
        if (format.FontSize > 0) range.Font.Size = format.FontSize;
        range.Font.Bold = (int)format.Bold;
        range.Font.Italic = (int)format.Italic;
        range.Font.Underline = format.Underline == EnumBoolean.True
            ? Interop.WdUnderline.wdUnderlineSingle
            : Interop.WdUnderline.wdUnderlineNone;
        range.Font.StrikeThrough = (int)format.Strikeout;

        if (format.Alignment != HeaderFooterParagraphAlignment.Select)
            range.ParagraphFormat.Alignment = (Interop.WdParagraphAlignment)(int)format.Alignment;
    }

    /// <inheritdoc />
    public void InsertHeaderFooterImage(EnumHeadersFooters insertType, HeaderFooterImages images)
    {
        ArgumentNullException.ThrowIfNull(images);
        foreach (Interop.Section section in Native.Sections)
        {
            Place(section, insertType, Interop.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                  images.OddPageImage, images.Alignment);
            Place(section, insertType, Interop.WdHeaderFooterIndex.wdHeaderFooterEvenPages,
                  images.EvenPageImage, images.Alignment);
            Place(section, insertType, Interop.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                  images.FirstPageImage, images.Alignment);
        }
    }

    private static void Place(Interop.Section section, EnumHeadersFooters kind,
                              Interop.WdHeaderFooterIndex index, string imagePath,
                              HeaderFooterParagraphAlignment alignment)
    {
        if (string.IsNullOrWhiteSpace(imagePath)) return;
        if (!File.Exists(imagePath))
            throw new FileNotFoundException($"The image '{imagePath}' does not exist.", imagePath);

        var parts = kind == EnumHeadersFooters.Header ? section.Headers : section.Footers;
        var range = parts[index].Range;
        range.InlineShapes.AddPicture(imagePath, LinkToFile: false, SaveWithDocument: true);

        if (alignment != HeaderFooterParagraphAlignment.Select)
            range.ParagraphFormat.Alignment = (Interop.WdParagraphAlignment)(int)alignment;
    }

    /// <inheritdoc />
    public void RemoveHeaderFooter(EnumHeadersFooters removeType)
    {
        foreach (Interop.Section section in Native.Sections)
        {
            var parts = removeType == EnumHeadersFooters.Header ? section.Headers : section.Footers;
            foreach (Interop.HeaderFooter part in parts) part.Range.Delete();
        }
    }

    /// <inheritdoc />
    public void ExtractHeaderFooterImages(EnumHeadersFooters readType, string saveFolder)
    {
        if (string.IsNullOrWhiteSpace(saveFolder))
            throw new ArgumentException("SaveFolder is required.", nameof(saveFolder));
        Directory.CreateDirectory(saveFolder);

        var index = 0;
        foreach (Interop.Section section in Native.Sections)
        {
            var parts = readType == EnumHeadersFooters.Header ? section.Headers : section.Footers;
            foreach (Interop.HeaderFooter part in parts)
            {
                foreach (Interop.InlineShape shape in part.Range.InlineShapes)
                {
                    index++;
                    shape.Range.CopyAsPicture();
                    SaveClipboardImage(
                        Path.Combine(saveFolder, $"{readType}{index}.png"), EnumFileExtension.Png);
                }
            }
        }
    }

    // ------------------------------------------------------------- readers

    /// <inheritdoc />
    public (string[] Array, DataTable Table) ReadByFont(EnumBoldItalicUnderline fontStyle)
    {
        bool Matches(Interop.Font font) => fontStyle switch
        {
            EnumBoldItalicUnderline.Bold => font.Bold != 0,
            EnumBoldItalicUnderline.Italic => font.Italic != 0,
            EnumBoldItalicUnderline.Underline => font.Underline != Interop.WdUnderline.wdUnderlineNone,
            _ => false,
        };

        return Collect(paragraph => Matches(paragraph.Range.Font), "ReadByFont");
    }

    /// <inheritdoc />
    public (string[] Array, DataTable Table) ReadByStyle(string paragraphStyle)
    {
        if (string.IsNullOrWhiteSpace(paragraphStyle))
            throw new ArgumentException("ParagraphStyle is required.", nameof(paragraphStyle));

        return Collect(
            paragraph => string.Equals(
                ((Interop.Style)paragraph.get_Style()).NameLocal,
                paragraphStyle,
                StringComparison.OrdinalIgnoreCase),
            "ReadByStyle");
    }

    private (string[] Array, DataTable Table) Collect(
        Func<Interop.Paragraph, bool> predicate, string tableName)
    {
        var matches = new List<string>();
        foreach (Interop.Paragraph paragraph in Native.Paragraphs)
        {
            if (!predicate(paragraph)) continue;
            var text = (paragraph.Range.Text ?? string.Empty).TrimEnd('\r');
            if (!string.IsNullOrWhiteSpace(text)) matches.Add(text);
        }

        var table = new DataTable(tableName);
        table.Columns.Add("Text", typeof(string));
        foreach (var match in matches) table.Rows.Add(match);
        return ([.. matches], table);
    }

    // -------------------------------------------------------------- tables

    /// <inheritdoc />
    public int TableCount() => Native.Tables.Count;

    /// <inheritdoc />
    public WordTableInfo TableInfo(int tableIndex)
    {
        var table = Table(tableIndex);
        return new WordTableInfo
        {
            TotalRows = table.Rows.Count,
            TotalColumns = table.Columns.Count,
            HasHeaderRow = table.ApplyStyleHeadingRows,
            HasBandedRows = table.ApplyStyleRowBands,
            HasBandedColumns = table.ApplyStyleColumnBands,
        };
    }

    /// <inheritdoc />
    public DataSet ReadAllTables(bool withHeader)
    {
        var set = new DataSet("WordTables");
        var index = 0;
        foreach (Interop.Table table in Native.Tables)
        {
            index++;
            var data = new DataTable($"Table{index}");
            var rows = table.Rows.Count;
            var columns = table.Columns.Count;

            for (var c = 1; c <= columns; c++)
            {
                var name = withHeader ? CellText(table, 1, c) : $"Column{c}";
                if (string.IsNullOrWhiteSpace(name)) name = $"Column{c}";
                // A Word table can repeat a heading; DataTable will not.
                while (data.Columns.Contains(name)) name += "_";
                data.Columns.Add(name, typeof(string));
            }

            for (var r = withHeader ? 2 : 1; r <= rows; r++)
            {
                var values = new object[columns];
                for (var c = 1; c <= columns; c++) values[c - 1] = CellText(table, r, c);
                data.Rows.Add(values);
            }
            set.Tables.Add(data);
        }
        return set;
    }

    /// <summary>
    /// Reads one cell, tolerating the ragged tables Word allows.
    /// </summary>
    /// <remarks>
    /// A merged or missing cell makes <c>Table.Cell</c> throw rather than return null,
    /// and a whole table read should not fail because one cell is spanned.
    /// </remarks>
    private static string CellText(Interop.Table table, int row, int column)
    {
        try
        {
            var text = table.Cell(row, column).Range.Text ?? string.Empty;
            // Word terminates every cell with \r\a.
            return text.TrimEnd('\a').TrimEnd('\r');
        }
        catch (COMException)
        {
            return string.Empty;
        }
    }

    /// <inheritdoc />
    public void InsertTable(EnumGoTo target, string bookmark, int pageOrLine, int rows, int columns)
    {
        if (rows <= 0 || columns <= 0)
            throw new ArgumentOutOfRangeException(nameof(rows), "Rows and columns must both be positive.");

        GoTo(target, bookmark, pageOrLine);
        Native.Tables.Add(_application.Selection.Range, rows, columns);
    }

    /// <inheritdoc />
    public void InsertDataTable(EnumGoTo target, string bookmark, int pageOrLine,
                                DataTable input, bool addHeader, string styleName)
    {
        ArgumentNullException.ThrowIfNull(input);

        GoTo(target, bookmark, pageOrLine);
        var rows = input.Rows.Count + (addHeader ? 1 : 0);
        var columns = input.Columns.Count;
        if (rows == 0 || columns == 0)
            throw new ArgumentException("InputTable has no rows or no columns.", nameof(input));

        var table = Native.Tables.Add(_application.Selection.Range, rows, columns);

        var offset = 1;
        if (addHeader)
        {
            for (var c = 0; c < columns; c++)
                table.Cell(1, c + 1).Range.Text = input.Columns[c].ColumnName;
            offset = 2;
        }

        for (var r = 0; r < input.Rows.Count; r++)
        {
            for (var c = 0; c < columns; c++)
                table.Cell(r + offset, c + 1).Range.Text = input.Rows[r][c]?.ToString() ?? string.Empty;
        }

        if (!string.IsNullOrWhiteSpace(styleName)) table.set_Style(styleName);
    }

    /// <inheritdoc />
    public void InsertTableRows(int tableIndex, int position, int count, float height, EnumInsertOption option)
    {
        var table = Table(tableIndex);
        if (count <= 0) throw new ArgumentOutOfRangeException(nameof(count), "NoRows must be positive.");

        for (var i = 0; i < count; i++)
        {
            var added = option == EnumInsertOption.AfterLast
                ? table.Rows.Add()
                : table.Rows.Add(table.Rows[Math.Clamp(position, 1, table.Rows.Count)]);
            if (height > 0) added.Height = height;
        }
    }

    /// <inheritdoc />
    public void InsertTableColumns(int tableIndex, int position, int count, float width, EnumInsertOption option)
    {
        var table = Table(tableIndex);
        if (count <= 0) throw new ArgumentOutOfRangeException(nameof(count), "NoColumns must be positive.");

        for (var i = 0; i < count; i++)
        {
            var added = option == EnumInsertOption.AfterLast
                ? table.Columns.Add()
                : table.Columns.Add(table.Columns[Math.Clamp(position, 1, table.Columns.Count)]);
            if (width > 0) added.Width = width;
        }
    }

    /// <inheritdoc />
    public void DeleteRow(int tableIndex, int rowIndex)
    {
        var table = Table(tableIndex);
        Bounds(rowIndex, table.Rows.Count, nameof(rowIndex));
        table.Rows[rowIndex].Delete();
    }

    /// <inheritdoc />
    public void DeleteColumn(int tableIndex, int columnIndex)
    {
        var table = Table(tableIndex);
        Bounds(columnIndex, table.Columns.Count, nameof(columnIndex));
        table.Columns[columnIndex].Delete();
    }

    /// <inheritdoc />
    public void DeleteTable(int tableIndex) => Table(tableIndex).Delete();

    /// <inheritdoc />
    public void RowHeight(int tableIndex, int[] rowIndexes, float height, EnumRowHeightRule rule)
    {
        var table = Table(tableIndex);
        // No indexes given means every row, which is the only reading that makes an
        // empty array useful rather than a silent no-op.
        var targets = rowIndexes is { Length: > 0 }
            ? rowIndexes
            : [.. Enumerable.Range(1, table.Rows.Count)];

        foreach (var index in targets)
        {
            Bounds(index, table.Rows.Count, nameof(rowIndexes));
            var row = table.Rows[index];
            row.HeightRule = (Interop.WdRowHeightRule)(int)rule;
            if (rule != EnumRowHeightRule.Auto && height > 0) row.Height = height;
        }
    }

    /// <inheritdoc />
    public void SetTableValue(int tableIndex, int rowIndex, int columnIndex, WordCellFormat format)
    {
        ArgumentNullException.ThrowIfNull(format);
        var table = Table(tableIndex);
        Bounds(rowIndex, table.Rows.Count, nameof(rowIndex));
        Bounds(columnIndex, table.Columns.Count, nameof(columnIndex));

        var cell = table.Cell(rowIndex, columnIndex);
        cell.Range.Text = format.TextValue;

        var font = cell.Range.Font;
        if (!string.IsNullOrWhiteSpace(format.FontName)) font.Name = format.FontName;
        if (format.FontSize > 0) font.Size = format.FontSize;
        Apply(format.Bold, value => font.Bold = value);
        Apply(format.Italic, value => font.Italic = value);
        Apply(format.Strikeout, value => font.StrikeThrough = value);
        Apply(format.Superscript, value => font.Superscript = value);
        Apply(format.Subscript, value => font.Subscript = value);
        Apply(format.Underline, value => font.Underline =
            value == 1 ? Interop.WdUnderline.wdUnderlineSingle : Interop.WdUnderline.wdUnderlineNone);

        if (format.VerticalAlignment != EnumCellVerticalAlignment.Select)
            cell.VerticalAlignment = (Interop.WdCellVerticalAlignment)(int)format.VerticalAlignment;
    }

    /// <summary>
    /// Applies a tri-state font flag, leaving the run alone on <c>Select</c>.
    /// </summary>
    /// <remarks>
    /// EnumSelectBoolean is Select=1, True=2, False=3, so it cannot be cast to a Word
    /// boolean directly: Select would read as true and False as a nonzero truth.
    /// </remarks>
    private static void Apply(EnumSelectBoolean flag, Action<int> set)
    {
        if (flag == EnumSelectBoolean.Select) return;
        set(flag == EnumSelectBoolean.True ? 1 : 0);
    }

    /// <inheritdoc />
    public void TableAutoFit(int tableIndex, EnumAutoFitBehavior autoFit) =>
        Table(tableIndex).AutoFitBehavior((Interop.WdAutoFitBehavior)(int)autoFit);

    /// <inheritdoc />
    public void TableStyle(int tableIndex, string styleName, WordTableStyleOptions options)
    {
        ArgumentNullException.ThrowIfNull(options);
        var table = Table(tableIndex);

        if (!string.IsNullOrWhiteSpace(styleName)) table.set_Style(styleName);
        table.ApplyStyleHeadingRows = options.HeaderRow;
        table.ApplyStyleLastRow = options.TotalRow;
        table.ApplyStyleFirstColumn = options.FirstColumn;
        table.ApplyStyleLastColumn = options.LastColumn;
        table.ApplyStyleRowBands = options.BandedRows;
        table.ApplyStyleColumnBands = options.BandedColumns;
    }

    /// <inheritdoc />
    public void CopyTableToClipboard(int tableIndex) => Table(tableIndex).Range.Copy();

    // ---------------------------------------------------------------- tools

    /// <inheritdoc />
    public void CloseAllWord()
    {
        foreach (Interop.Document open in _application.Documents) open.Close(SaveChanges: false);
        _closed = true;
        Quit();
    }

    // ------------------------------------------------------------- plumbing

    private Interop.Table Table(int tableIndex)
    {
        var count = Native.Tables.Count;
        if (count == 0) throw new InvalidOperationException("The document has no tables.");
        Bounds(tableIndex, count, nameof(tableIndex));
        return Native.Tables[tableIndex];
    }

    private static void Bounds(int index, int count, string name)
    {
        if (index < 1 || index > count)
            throw new ArgumentOutOfRangeException(name, index, $"Must be between 1 and {count}.");
    }

    private void GoTo(EnumGoTo target, string bookmark, int pageOrLine)
    {
        switch (target)
        {
            case EnumGoTo.Bookmark:
                if (string.IsNullOrWhiteSpace(bookmark))
                    throw new ArgumentException("BookMark is required when GoTo is Bookmark.", nameof(bookmark));
                if (!Native.Bookmarks.Exists(bookmark))
                    throw new ArgumentException($"The document has no bookmark named '{bookmark}'.", nameof(bookmark));
                Native.Bookmarks[bookmark].Select();
                break;

            case EnumGoTo.Page:
            case EnumGoTo.Line:
                _application.Selection.GoTo(
                    What: target == EnumGoTo.Page
                        ? Interop.WdGoToItem.wdGoToPage
                        : Interop.WdGoToItem.wdGoToLine,
                    Which: Interop.WdGoToDirection.wdGoToAbsolute,
                    Count: Math.Max(pageOrLine, 1));
                break;

            default:
                throw new ArgumentOutOfRangeException(nameof(target), target, "Unknown GoTo target.");
        }
    }

    /// <summary>Writes whatever CopyAsPicture put on the clipboard to a file.</summary>
    private static void SaveClipboardImage(string path, EnumFileExtension extension)
    {
        // Deliberately not implemented through System.Windows.Forms: pulling WinForms in
        // for a clipboard read would add a WPF-sized dependency to a document package.
        // Word's own export is used instead where a caller needs image files.
        throw new NotSupportedException(
            $"Extracting images to '{path}' as {extension} is not implemented. " +
            "See docs/REVIVAL.md: the published activity round-tripped images through the " +
            "Windows clipboard, which this reimplementation does not do.");
    }

    private void Quit()
    {
        try { _application.Quit(SaveChanges: false); }
        catch (COMException) { /* Word was already gone. */ }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_closed) return;
        _closed = true;
        try { Native?.Close(SaveChanges: true); }
        catch (COMException) { /* nothing to save into */ }
        Quit();
    }
}
