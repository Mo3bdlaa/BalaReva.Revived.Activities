using System.Activities;
using System.Activities.Statements;
using System.Data;
using BalaReva.Word.Utilities;

namespace BalaReva.Word.Tests;

/// <summary>
/// A stand-in Word session that records what it was asked to do.
/// </summary>
/// <remarks>
/// No CI agent has Word installed, so this is the only way any of these activities get
/// exercised. It covers argument handling, which call is made with which values, and
/// how results are mapped back onto output arguments. The COM half is unreachable from
/// a test.
/// </remarks>
public sealed class FakeWordService : IWordService, IWordDocument
{
    public List<string> Calls { get; } = [];

    public bool Disposed { get; private set; }

    public object? MacroResult { get; set; }

    public string[] HeaderFooterText { get; set; } = ["Page header"];

    public int Tables { get; set; } = 2;

    public WordStatisticsResult Stats { get; set; } = new() { Pages = 3, WordCount = 120 };

    public WordTableInfo Info { get; set; } = new() { TotalRows = 4, TotalColumns = 3 };

    public WordCellFormat? LastCellFormat { get; private set; }

    public WordTableStyleOptions? LastStyleOptions { get; private set; }

    public HeaderFooterText? LastHeaderFooter { get; private set; }

    public Exception? Throw { get; set; }

    private T Record<T>(string call, T result)
    {
        Calls.Add(call);
        if (Throw is not null) throw Throw;
        return result;
    }

    // IWordService

    public IWordDocument Open(string filePath, string openPassword, string modifyPassword)
    {
        Target = new WordObject
        {
            FilePath = filePath,
            Password = openPassword,
            ModiPassword = modifyPassword,
        };

        // Deliberately does not honour Throw: that flag stands for a document operation
        // failing, which is what ContinueOnError and the scope's fault path are about.
        // Failing the open instead would mean neither ever runs.
        Calls.Add($"Open({filePath})");
        return this;
    }

    public void MergeDocuments(string wordFile, string openPassword, string modifyPassword,
                               string[] sourceFiles) =>
        Record($"Merge({wordFile},[{string.Join(",", sourceFiles)}])", 0);

    public void WordToPdf(string wordFile, string openPassword, string modifyPassword,
                          string pdfFile, int startPage, int endPage) =>
        Record($"WordToPdf({wordFile},{pdfFile},{startPage}-{endPage})", 0);

    // IWordDocument

    public WordObject Target { get; private set; } = new();

    public void ChangePassword(string newOpenPassword, string newModifyPassword) =>
        Record($"ChangePassword({newOpenPassword},{newModifyPassword})", 0);

    public void CreateDocument() => Record("CreateDocument", 0);

    public void SaveAs(string newFileName, EnumSaveAs format) =>
        Record($"SaveAs({newFileName},{format})", 0);

    public void Print(string printerName, int copies, EnumOrientation orientation) =>
        Record($"Print({printerName},{copies},{orientation})", 0);

    public void Select(EnumGoTo target, string bookmark, int pageOrLine, int rows, int columns) =>
        Record($"Select({target},{bookmark},{pageOrLine},{rows}x{columns})", 0);

    public void Paste(EnumGoTo target, string bookmark, int pageOrLine) =>
        Record($"Paste({target},{bookmark},{pageOrLine})", 0);

    public void AddPageBreak() => Record("AddPageBreak", 0);

    public void RemoveDisplayLineNumber() => Record("RemoveDisplayLineNumber", 0);

    public void FindReplace(string findText, string replaceText, EnumFindReplace findOption,
                            EnumReplaceOption replaceOption, bool matchCase) =>
        Record($"FindReplace({findText}->{replaceText},{findOption},{replaceOption},case={matchCase})", 0);

    public object? ExecuteMacro(string macroName, object[]? parameters) =>
        Record($"ExecuteMacro({macroName},[{string.Join(",", parameters ?? [])}])", MacroResult);

    public WordStatisticsResult Statistics() => Record("Statistics", Stats);

    public void ExtractImages(string imageFolder, EnumFileExtension extension) =>
        Record($"ExtractImages({imageFolder},{extension})", 0);

    public string[] ReadHeaderFooter(EnumHeadersFooters readType) =>
        Record($"ReadHeaderFooter({readType})", HeaderFooterText);

    public void InsertHeaderFooter(EnumHeadersFooters insertType, BalaReva.Word.HeaderFooterText text)
    {
        LastHeaderFooter = text;
        Record($"InsertHeaderFooter({insertType})", 0);
    }

    public void InsertHeaderFooterImage(EnumHeadersFooters insertType, HeaderFooterImages images) =>
        Record($"InsertHeaderFooterImage({insertType},{images.OddPageImage})", 0);

    public void RemoveHeaderFooter(EnumHeadersFooters removeType) =>
        Record($"RemoveHeaderFooter({removeType})", 0);

    public void ExtractHeaderFooterImages(EnumHeadersFooters readType, string saveFolder) =>
        Record($"ExtractHeaderFooterImages({readType},{saveFolder})", 0);

    public (string[] Array, DataTable Table) ReadByFont(EnumBoldItalicUnderline fontStyle) =>
        Record($"ReadByFont({fontStyle})", (new[] { "bold text" }, Single("bold text")));

    public (string[] Array, DataTable Table) ReadByStyle(string paragraphStyle) =>
        Record($"ReadByStyle({paragraphStyle})", (new[] { "heading" }, Single("heading")));

    public int TableCount() => Record("TableCount", Tables);

    public WordTableInfo TableInfo(int tableIndex) => Record($"TableInfo({tableIndex})", Info);

    public DataSet ReadAllTables(bool withHeader) =>
        Record($"ReadAllTables(header={withHeader})", new DataSet("WordTables"));

    public void InsertTable(EnumGoTo target, string bookmark, int pageOrLine, int rows, int columns) =>
        Record($"InsertTable({target},{bookmark},{pageOrLine},{rows}x{columns})", 0);

    public void InsertDataTable(EnumGoTo target, string bookmark, int pageOrLine,
                                DataTable input, bool addHeader, string styleName) =>
        Record($"InsertDataTable({target},rows={input.Rows.Count},header={addHeader},{styleName})", 0);

    public void InsertTableRows(int tableIndex, int position, int count, float height, EnumInsertOption option) =>
        Record($"InsertTableRows({tableIndex},{position},{count},{height},{option})", 0);

    public void InsertTableColumns(int tableIndex, int position, int count, float width, EnumInsertOption option) =>
        Record($"InsertTableColumns({tableIndex},{position},{count},{width},{option})", 0);

    public void DeleteRow(int tableIndex, int rowIndex) => Record($"DeleteRow({tableIndex},{rowIndex})", 0);

    public void DeleteColumn(int tableIndex, int columnIndex) =>
        Record($"DeleteColumn({tableIndex},{columnIndex})", 0);

    public void DeleteTable(int tableIndex) => Record($"DeleteTable({tableIndex})", 0);

    public void RowHeight(int tableIndex, int[] rowIndexes, float height, EnumRowHeightRule rule) =>
        Record($"RowHeight({tableIndex},[{string.Join(",", rowIndexes)}],{height},{rule})", 0);

    public void SetTableValue(int tableIndex, int rowIndex, int columnIndex, WordCellFormat format)
    {
        LastCellFormat = format;
        Record($"SetTableValue({tableIndex},{rowIndex},{columnIndex},{format.TextValue})", 0);
    }

    public void TableAutoFit(int tableIndex, EnumAutoFitBehavior autoFit) =>
        Record($"TableAutoFit({tableIndex},{autoFit})", 0);

    public void TableStyle(int tableIndex, string styleName, WordTableStyleOptions options)
    {
        LastStyleOptions = options;
        Record($"TableStyle({tableIndex},{styleName})", 0);
    }

    public void CopyTableToClipboard(int tableIndex) => Record($"CopyTableToClipboard({tableIndex})", 0);

    public void CloseAllWord() => Record("CloseAllWord", 0);

    public void Dispose() => Disposed = true;

    private static DataTable Single(string value)
    {
        var table = new DataTable("Result");
        table.Columns.Add("Text", typeof(string));
        table.Rows.Add(value);
        return table;
    }
}

/// <summary>Runs one in-scope activity inside a real WordScope over a stand-in session.</summary>
/// <remarks>
/// A child runs nested inside the scope, so its output arguments never reach the
/// workflow root and <c>Invoke</c> returns nothing for them. Each one has to be routed
/// into a variable and read back after the scope closes, which is what
/// <see cref="ScopeRun"/> is for.
/// </remarks>
internal static class Harness
{
    public static IDictionary<string, object> Run(
        BaseNativeChild child, FakeWordService service, string filePath = @"C:\docs\test.docx") =>
        new ScopeRun().Invoke(child, service, filePath).Values;

    /// <summary>Runs a standalone activity, which needs no scope.</summary>
    public static void RunStandalone(BaseWord activity, FakeWordService service)
    {
        activity.ContinueOnError ??= new InArgument<bool>(false);
        var invoker = new WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        invoker.Invoke();
    }
}


/// <summary>Whatever the captured output arguments held when the workflow finished.</summary>
internal sealed class Outputs
{
    public Dictionary<string, object> Values { get; } = [];
}

/// <summary>Copies a workflow variable into <see cref="Outputs"/> once the scope is done.</summary>
internal sealed class CaptureOutput<T>(Outputs sink, string name) : CodeActivity
{
    public InArgument<T> Value { get; set; } = null!;

    protected override void Execute(CodeActivityContext context) =>
        sink.Values[name] = context.GetValue(Value)!;
}

/// <summary>
/// Builds a workflow that runs one activity inside a WordScope and reads its outputs.
/// </summary>
internal sealed class ScopeRun
{
    private readonly List<Variable> _variables = [];
    private readonly List<Activity> _captures = [];
    private readonly Outputs _outputs = new();

    /// <summary>
    /// Declares a variable for one of the child's output arguments and arranges for it
    /// to be read back under <paramref name="name"/>.
    /// </summary>
    public OutArgument<T> Capture<T>(string name)
    {
        var variable = new Variable<T>();
        _variables.Add(variable);
        _captures.Add(new CaptureOutput<T>(_outputs, name) { Value = new InArgument<T>(variable) });
        return new OutArgument<T>(variable);
    }

    public Outputs Invoke(
        BaseNativeChild child, FakeWordService service, string filePath = @"C:\docs\test.docx")
    {
        child.ContinueOnError ??= new InArgument<bool>(false);
        child.Delay ??= new InArgument<double>(0);
        child.ExecutionResult ??= new OutArgument<bool>();

        var scope = new WordScope
        {
            FilePath = new InArgument<string>(filePath),
            Body = new ActivityAction<WordObject>
            {
                Argument = new DelegateInArgument<WordObject> { Name = "WordDocument" },
                Handler = child,
            },
        };

        var root = new Sequence { Variables = { }, Activities = { scope } };
        foreach (var variable in _variables) root.Variables.Add(variable);
        foreach (var capture in _captures) root.Activities.Add(capture);

        var invoker = new WorkflowInvoker(root);
        invoker.Extensions.Add(service);
        invoker.Invoke();
        return _outputs;
    }
}
