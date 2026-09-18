using System.Activities;
using System.Activities.Statements;
using System.Data;
using BalaReva.Easy.PowerPoint.Utilities;
using BalaReva.EasyPowerPoint.Base;
using BalaReva.EasyPowerPoint.Scope.Main;
using BalaReva.EasyPowerPoint.Utilities;

namespace BalaReva.EasyPowerPoint.Tests;

/// <summary>
/// A stand-in presentation that records what it was asked to do.
/// </summary>
/// <remarks>
/// No CI agent has PowerPoint installed, so this is the only way these activities get
/// exercised. It covers argument handling, which call is made with which values, and
/// how results are mapped back onto output arguments.
/// </remarks>
public sealed class FakePowerPointService : IPowerPointService, IPowerPointPresentation
{
    public List<string> Calls { get; } = [];

    public bool Disposed { get; private set; }

    public bool MacrosWereEnabled { get; private set; }

    public bool AlertsWereOn { get; private set; }

    public int Slides { get; set; } = 3;

    public object? MacroResult { get; set; }

    public string[] TableNames { get; set; } = ["Sales", "Costs"];

    public string CellText { get; set; } = "42";

    public TextBoxRequest? LastTextBox { get; private set; }

    public PictureRequest? LastPicture { get; private set; }

    public AddTableRequest? LastTable { get; private set; }

    public TableStyleOptions? LastStyle { get; private set; }

    public TextStyleRequest? LastFont { get; private set; }

    public TableRef? LastTableRef { get; private set; }

    public Exception? Throw { get; set; }

    private T Record<T>(string call, T result)
    {
        Calls.Add(call);
        if (Throw is not null) throw Throw;
        return result;
    }

    // IPowerPointService

    public IPowerPointPresentation Open(string filePath, string openPassword, string modifyPassword,
                                        bool displayAlerts, bool macrosEnabled)
    {
        FilePath = filePath;
        MacrosWereEnabled = macrosEnabled;
        AlertsWereOn = displayAlerts;
        // Deliberately ignores Throw: that flag stands for an operation failing, and
        // failing the open would mean ContinueOnError never got to do its job.
        Calls.Add($"Open({filePath})");
        return this;
    }

    // IPowerPointPresentation

    public string FilePath { get; private set; } = string.Empty;

    public int SlideCount() => Record("SlideCount", Slides);

    public void NewSlide(int slideIndex) => Record($"NewSlide({slideIndex})", 0);

    public void DeleteSlide(int slideIndex) => Record($"DeleteSlide({slideIndex})", 0);

    public void DuplicateSlide(int slideIndex) => Record($"DuplicateSlide({slideIndex})", 0);

    public void SlideCopy(int slideIndex) => Record($"SlideCopy({slideIndex})", 0);

    public void SlidePaste(int slideIndex) => Record($"SlidePaste({slideIndex})", 0);

    public string SlideExtractor(int slideIndex) =>
        Record($"SlideExtractor({slideIndex})", $@"C:\out\slide{slideIndex}.pptx");

    public void HideUnhideSlide(int slideIndex, HideUnhideEnum slideShow) =>
        Record($"HideUnhideSlide({slideIndex},{slideShow})", 0);

    public void SlideTransitions(int slideIndex, EntryEffectEnum effect, bool onMouseClick, float duration) =>
        Record($"SlideTransitions({slideIndex},{effect},{onMouseClick},{duration})", 0);

    public void Merge(string sourcePpt, int startSlideIndex, int endSlideIndex, int slideAfter) =>
        Record($"Merge({sourcePpt},{startSlideIndex}-{endSlideIndex},after={slideAfter})", 0);

    public (string[] Array, string Text) ReadText(int[] slideIndexes, bool addSlideIndex, bool omitEmptyLine) =>
        Record($"ReadText([{string.Join(",", slideIndexes)}],{addSlideIndex},{omitEmptyLine})",
               (new[] { "one", "two" }, "one\ntwo"));

    public (string[] Array, DataTable Table) FindText(int[] slideIndexes, string find,
                                                      bool matchCase, bool wholeWord) =>
        Record($"FindText([{string.Join(",", slideIndexes)}],{find},{matchCase},{wholeWord})",
               (new[] { find }, new DataTable("FindText")));

    public void FindReplace(int[] slideIndexes, string find, string replace,
                            bool matchCase, bool wholeWord, bool firstOccurrence) =>
        Record($"FindReplace([{string.Join(",", slideIndexes)}],{find}->{replace},"
               + $"{matchCase},{wholeWord},first={firstOccurrence})", 0);

    public void InsertTextBox(int slideIndex, TextBoxRequest request)
    {
        LastTextBox = request;
        Record($"InsertTextBox({slideIndex},{request.Text})", 0);
    }

    public void TextShapeEdit(int slideIndex, int textIndex, TextStyleRequest style)
    {
        LastFont = style;
        Record($"TextShapeEdit({slideIndex},{textIndex})", 0);
    }

    public int TextShapeCount(int slideIndex) => Record($"TextShapeCount({slideIndex})", 7);

    public DataTable ExtractHyperLinks(int[] slideIndexes) =>
        Record($"ExtractHyperLinks([{string.Join(",", slideIndexes)}])", new DataTable("HyperLinks"));

    public void InsertPicture(int slideIndex, PictureRequest request)
    {
        LastPicture = request;
        Record($"InsertPicture({slideIndex},{request.ImagePath})", 0);
    }

    public void DeleteImage(int slideIndex, int imageIndex) =>
        Record($"DeleteImage({slideIndex},{imageIndex})", 0);

    public int ImageShapeCount(int slideIndex) => Record($"ImageShapeCount({slideIndex})", 5);

    public void ImageExtractor(int slideIndex, string imageDirectory, ImageFileFormatEnum format) =>
        Record($"ImageExtractor({slideIndex},{imageDirectory},{format})", 0);

    public void PasteClipboard(int slideIndex, double left, double top, double width, double height) =>
        Record($"PasteClipboard({slideIndex},{left},{top},{width},{height})", 0);

    public int ChartShapeCount(int slideIndex) => Record($"ChartShapeCount({slideIndex})", 2);

    public void ChartDelete(int slideIndex, int chartIndex) =>
        Record($"ChartDelete({slideIndex},{chartIndex})", 0);

    public void ChartCopyToClipboard(int slideIndex, int chartIndex) =>
        Record($"ChartCopyToClipboard({slideIndex},{chartIndex})", 0);

    public void ChartFormat(int slideIndex, int chartIndex, double left, double top,
                            double width, double height) =>
        Record($"ChartFormat({slideIndex},{chartIndex},{left},{top},{width},{height})", 0);

    public void ChartImageExtract(int slideIndex, string imageFolder, ImageFileFormatEnum format) =>
        Record($"ChartImageExtract({slideIndex},{imageFolder},{format})", 0);

    public void RefreshData(int slideIndex) => Record($"RefreshData({slideIndex})", 0);

    public void UpdateLinks() => Record("UpdateLinks", 0);

    public void CommentsAdd(int slideIndex, string author, string commentText, float left, float top) =>
        Record($"CommentsAdd({slideIndex},{author},{commentText},{left},{top})", 0);

    public void CommentsDelete(int slideIndex) => Record($"CommentsDelete({slideIndex})", 0);

    public (string Text, DataTable Table) CommentsRead(int slideIndex, bool includeReplies) =>
        Record($"CommentsRead({slideIndex},replies={includeReplies})",
               ("alice: nice", new DataTable("Comments")));

    public void AddTable(int slideIndex, AddTableRequest request)
    {
        LastTable = request;
        Record($"AddTable({slideIndex},{request.TableName},rows={request.InputTable.Rows.Count})", 0);
    }

    public void DeleteTable(TableRef table) => Track("DeleteTable", table);

    public void ClearTable(TableRef table, bool leaveFirstRow) =>
        Track($"ClearTable(leaveFirst={leaveFirstRow})", table);

    public void AppendTable(TableRef table, string[] values) =>
        Track($"AppendTable([{string.Join(",", values)}])", table);

    public void EditTable(TableRef table, int rowIndex, int columnIndex, string cellValue) =>
        Track($"EditTable({rowIndex},{columnIndex},{cellValue})", table);

    public string GetRowItem(TableRef table, int rowIndex, int columnIndex)
    {
        LastTableRef = table;
        return Record($"GetRowItem({Describe(table)},{rowIndex},{columnIndex})", CellText);
    }

    public void DeleteRow(TableRef table, int rowIndex) => Track($"DeleteRow({rowIndex})", table);

    public void DeleteColumn(TableRef table, int columnIndex) =>
        Track($"DeleteColumn({columnIndex})", table);

    public void ResizeTable(TableRef table, double left, double top, double width, double height) =>
        Track($"ResizeTable({left},{top},{width},{height})", table);

    public void FontOption(TableRef table, TextStyleRequest style)
    {
        LastFont = style;
        Track("FontOption", table);
    }

    public void StyleOption(TableRef table, TableStyleOptions options)
    {
        LastStyle = options;
        Track("StyleOption", table);
    }

    public void TableCopyToClipboard(TableRef table) => Track("TableCopyToClipboard", table);

    public DataSet ExtractTables(int slideIndex) =>
        Record($"ExtractTables({slideIndex})", new DataSet("SlideTables"));

    public string[] GetTableNames(int slideIndex) =>
        Record($"GetTableNames({slideIndex})", TableNames);

    public void DataTransformer(int[] slideIndexes, Dictionary<string, string> replacements) =>
        Record($"DataTransformer([{string.Join(",", slideIndexes)}],{replacements.Count} pairs)", 0);

    public void ExportTableToExcel(TableRef table, string excelFile, string sheetName, string startCell) =>
        Track($"ExportTableToExcel({excelFile},{sheetName},{startCell})", table);

    public void ImportDataFromExcel(int slideIndex, string excelFile, string sheetName, string cellRange) =>
        Record($"ImportDataFromExcel({slideIndex},{excelFile},{sheetName},{cellRange})", 0);

    public void ExportPdf(string filePath, FixedFormatIntentEnum formatType) =>
        Record($"ExportPdf({filePath},{formatType})", 0);

    public void SaveAs(string saveAsFile, SaveAsEnum format) =>
        Record($"SaveAs({saveAsFile},{format})", 0);

    public void Print(int numberOfCopies, PrintColorTypeEnum colorType,
                      bool printComments, bool printHiddenSlides) =>
        Record($"Print({numberOfCopies},{colorType},{printComments},{printHiddenSlides})", 0);

    public void RemoveDocumentInformation(RemoveDocInfoTypeEnum docInfoType) =>
        Record($"RemoveDocumentInformation({docInfoType})", 0);

    public object? ExecuteMacro(string macroName, object[]? arguments) =>
        Record($"ExecuteMacro({macroName},[{string.Join(",", arguments ?? [])}])", MacroResult);

    public void Dispose() => Disposed = true;

    private void Track(string call, TableRef table)
    {
        LastTableRef = table;
        Record($"{call}@{Describe(table)}", 0);
    }

    private static string Describe(TableRef table) =>
        string.IsNullOrWhiteSpace(table.TableName)
            ? $"slide{table.SlideIndex}/table{table.TableIndex}"
            : $"slide{table.SlideIndex}/{table.TableName}";
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
/// Builds a workflow that runs one activity inside a PowerPointScope and reads its
/// outputs back, which a nested activity cannot return on its own.
/// </summary>
internal sealed class ScopeRun
{
    private readonly List<Variable> _variables = [];
    private readonly List<Activity> _captures = [];
    private readonly Outputs _outputs = new();

    public OutArgument<T> Capture<T>(string name)
    {
        var variable = new Variable<T>();
        _variables.Add(variable);
        _captures.Add(new CaptureOutput<T>(_outputs, name) { Value = new InArgument<T>(variable) });
        return new OutArgument<T>(variable);
    }

    public Outputs Invoke(BaseNativeChild child, FakePowerPointService service,
                          string filePath = @"C:\decks\deck.pptx",
                          EnableDisableEnum macros = EnableDisableEnum.Disable)
    {
        child.ContinueOnError ??= new InArgument<bool>(false);
        child.Delay ??= new InArgument<double>(0);
        child.ExecutionResult ??= new OutArgument<bool>();

        var scope = new PowerPointScope
        {
            FilePath = new InArgument<string>(filePath),
            MacroSettings = macros,
            Body = new ActivityAction<string>
            {
                Argument = new DelegateInArgument<string> { Name = "FilePath" },
                Handler = child,
            },
        };

        var root = new Sequence { Activities = { scope } };
        foreach (var variable in _variables) root.Variables.Add(variable);
        foreach (var capture in _captures) root.Activities.Add(capture);

        var invoker = new WorkflowInvoker(root);
        invoker.Extensions.Add(service);
        invoker.Invoke();
        return _outputs;
    }
}

/// <summary>Runs one in-scope activity and returns the captured outputs.</summary>
internal static class Harness
{
    public static IDictionary<string, object> Run(
        BaseNativeChild child, FakePowerPointService service,
        string filePath = @"C:\decks\deck.pptx") =>
        new ScopeRun().Invoke(child, service, filePath).Values;
}
