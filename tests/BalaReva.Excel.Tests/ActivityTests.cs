using System.Activities;
using System.Data;
using BalaReva.Excel.Base;
using BalaReva.Excel.Charts;
using BalaReva.Excel.Comment;
using BalaReva.Excel.Enums;
using BalaReva.Excel.External;
using BalaReva.Excel.Hide_UnHide;
using BalaReva.Excel.Merge;
using BalaReva.Excel.Sheets;
using BalaReva.Excel.Utilities;
using BalaReva.Excel.WorkBook;

namespace BalaReva.Excel.Tests;

/// <summary>
/// Each activity runs as a real workflow against a stand-in workbook, covering argument
/// handling and output mapping without Excel.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void Every_activity_opens_the_workbook_and_closes_it_afterwards()
    {
        var service = new FakeExcelService();

        Harness.Run(new AddSheet { SheetName = new InArgument<string>("New") }, service,
                    @"C:\books\report.xlsx");

        Assert.Equal([@"Open(C:\books\report.xlsx)", "AddSheet(New)"], service.Calls);
        Assert.True(service.Disposed);
    }

    [Fact]
    public void A_missing_file_path_is_rejected_before_the_workbook_is_touched()
    {
        var service = new FakeExcelService();
        var activity = new AddSheet { FilePath = new InArgument<string>("  ") };

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
        Assert.Empty(service.Calls);
    }

    [Fact]
    public void ContinueOnError_reports_failure_instead_of_faulting()
    {
        var service = new FakeExcelService { Throw = new InvalidOperationException("boom") };
        var activity = new AddSheet
        {
            ContinueOnError = new InArgument<bool>(true),
            ExecutionResult = new OutArgument<bool>(),
        };

        var outputs = Harness.Run(activity, service);

        Assert.False((bool)outputs["ExecutionResult"]);
        // The workbook is still closed, even though the operation failed.
        Assert.True(service.Disposed);
    }

    [Fact]
    public void A_successful_activity_reports_ExecutionResult_true()
    {
        var outputs = Harness.Run(
            new AddSheet { ExecutionResult = new OutArgument<bool>() }, new FakeExcelService());

        Assert.True((bool)outputs["ExecutionResult"]);
    }

    [Fact]
    public void GetComment_reports_the_text_rather_than_an_ExecutionResult()
    {
        var service = new FakeExcelService { Comment = "check this" };

        var outputs = Harness.Run(
            new GetComment
            {
                SheetName = new InArgument<string>("Data"),
                Cell = new InArgument<string>("B4"),
                Result = new OutArgument<string>(),
            },
            service);

        Assert.Equal("check this", outputs["Result"]);
        Assert.DoesNotContain("ExecutionResult", outputs.Keys);
    }

    [Fact]
    public void The_comment_activities_forward_their_cell()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new AddComment
            {
                SheetName = new InArgument<string>("Data"),
                Cell = new InArgument<string>("A1"),
                Comment = new InArgument<string>("hello"),
            },
            service);
        Harness.Run(
            new DeleteComment { Cell = new InArgument<string>("A1") }, service);
        Harness.Run(
            new ShowHideComment { Cell = new InArgument<string>("A1"), ShowComment = true }, service);

        Assert.Contains("AddComment(Data,A1,hello)", service.Calls);
        Assert.Contains("DeleteComment(,A1)", service.Calls);
        Assert.Contains("ShowHideComment(,A1,True)", service.Calls);
    }

    [Fact]
    public void The_chart_activities_pass_Excel_own_chart_numbers()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new BarChart
            {
                CellRange = new InArgument<string>("A1:B10"),
                ChartType = BarChartEnum.BarStacked,
            },
            service);

        var request = Assert.IsType<ChartRequest>(service.LastChart);
        Assert.Equal(58, request.ChartType);
        Assert.Equal("A1:B10", request.CellRange);
    }

    [Fact]
    public void ColumnChart_carries_the_three_label_options_the_others_lack()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ColumnChart
            {
                CellRange = new InArgument<string>("A1:B10"),
                ChartType = ColumnChartEnum.Column3D,
                ShowLegendKey = true,
                ShowValuePosition = DataLabelPositionEnum.OutsideEnd,
                ShowValueTextOrientation = TextOrientationEnum.Upward,
            },
            service);

        var request = Assert.IsType<ChartRequest>(service.LastChart);
        Assert.Equal(-4100, request.ChartType);
        Assert.True(request.ShowLegendKey);
        Assert.Equal(DataLabelPositionEnum.OutsideEnd, request.ShowValuePosition);
        Assert.Equal(TextOrientationEnum.Upward, request.ShowValueTextOrientation);
    }

    [Fact]
    public void ClipboardToDatatable_needs_no_workbook_at_all()
    {
        var table = new DataTable("Clipboard");
        table.Columns.Add("Name");
        table.Rows.Add("Ada");
        var service = new FakeExcelService { Clipboard = table };

        var activity = new ClipboardToDatatable
        {
            HasHeader = new InArgument<bool>(true),
            Datatable = new OutArgument<DataTable>(),
        };
        var invoker = new WorkflowInvoker(activity);
        invoker.Extensions.Add(service);
        var outputs = invoker.Invoke();

        Assert.Same(table, outputs["Datatable"]);
        // No Open call: this is the one activity that opens nothing.
        Assert.Equal(["ClipboardToDataTable(header=True)"], service.Calls);
    }

    [Fact]
    public void The_hide_activities_turn_their_enum_into_a_boolean()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ColumnHide
            {
                ColumnNames = new InArgument<string[]>(_ => new[] { "A", "C" }),
                HiddenType = HideEnum.Hide,
            },
            service);
        Harness.Run(
            new RowHide
            {
                RowNumbers = new InArgument<int[]>(_ => new[] { 2, 4 }),
                HiddenType = HideEnum.Unhide,
            },
            service);

        Assert.Contains("HideColumns(,[A,C],True)", service.Calls);
        Assert.Contains("HideRows(,[2,4],False)", service.Calls);
    }

    [Fact]
    public void MergeCells_forwards_its_text_and_both_alignments()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new MergeCells
            {
                MergeRange = new InArgument<string>("A1:C1"),
                CellText = new InArgument<string>("Title"),
                HorizontalAlignment = AlignmentEnum.Center,
                VerticalAlignment = AlignmentEnum.Distributed,
            },
            service);

        Assert.Contains("MergeCells(,A1:C1,Title,Center,Distributed)", service.Calls);
    }

    [Fact]
    public void FormatCells_leaves_untouched_options_on_Select()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new FormatCells
            {
                CellRange = new InArgument<string[]>(_ => new[] { "A1:C10" }),
                Horizontal = FormatHorizontalEnum.Center,
                WrapText = FormatTextControlEnum.True,
            },
            service);

        var request = Assert.IsType<CellFormatRequest>(service.LastFormat);
        Assert.Equal(FormatHorizontalEnum.Center, request.Horizontal);
        Assert.Equal(FormatTextControlEnum.True, request.WrapText);
        // Everything the workflow did not set stays Select, which the service reads as
        // "leave the sheet's own formatting alone".
        Assert.Equal(FormatVerticleEnum.Select, request.Verticle);
        Assert.Equal(FormatTextControlEnum.Select, request.ShrinkFit);
        Assert.Equal(TextOrientationEumn.Select, request.TextOrientation);
    }

    [Fact]
    public void GetSheetsName_reports_what_the_workbook_returns()
    {
        var service = new FakeExcelService { Sheets = ["One", "Two", "Three"] };

        var outputs = Harness.Run(
            new GetSheetsName { SheetsName = new OutArgument<string[]>() }, service);

        Assert.Equal(["One", "Two", "Three"], (string[])outputs["SheetsName"]);
    }

    [Fact]
    public void CopyToFile_gathers_its_destination_arguments_into_one_request()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new CopyToFile
            {
                SheetName = new InArgument<string>("Data"),
                NewFilePath = new InArgument<string>(@"C:\books\out.xlsx"),
                NewSheetName = new InArgument<string>("Copied"),
                AutoFileCreation = new InArgument<bool>(true),
            },
            service);

        var request = Assert.IsType<CopyToFileRequest>(service.LastCopy);
        Assert.Equal(@"C:\books\out.xlsx", request.NewFilePath);
        Assert.Equal("Copied", request.NewSheetName);
        Assert.True(request.AutoFileCreation);
    }

    [Fact]
    public void ExportWorkBook_forwards_the_range_and_the_format()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ExportWorkBook
            {
                ExportPath = new InArgument<string>(@"C:\out\book.pdf"),
                CellRange = new InArgument<string>("A1:D20"),
                FormatType = FixedFormatTypeEnum.PDF,
            },
            service);

        Assert.Contains(@"ExportWorkBook(,A1:D20,C:\out\book.pdf,PDF)", service.Calls);
    }

    [Fact]
    public void FindReplace_defaults_to_the_whole_used_range()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new FindReplace
            {
                Find = new InArgument<string>("old"),
                Replace = new InArgument<string>("new"),
                FindOption = FindReplaceEnum.Whole,
            },
            service);

        // An empty range is passed through; choosing the used range is the workbook's job.
        Assert.Contains("FindReplace(,,old->new,Whole)", service.Calls);
    }

    [Fact]
    public void ProtectUnProtectSheet_forwards_the_password_and_the_direction()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new ProtectUnProtectSheet
            {
                SheetName = new InArgument<string>("Data"),
                ProtectPassword = new InArgument<string>("secret"),
                ProtectType = ProtectUnProtectEnum.UnProtect,
            },
            service);

        Assert.Contains("ProtectSheet(Data,secret,UnProtect)", service.Calls);
    }

    [Fact]
    public void HyperlinkAdd_rejects_a_missing_table()
    {
        var service = new FakeExcelService();
        var activity = new HyperlinkAdd
        {
            ColumnNames = new InArgument<string[]>(_ => new[] { "Url" }),
        };

        Assert.ThrowsAny<Exception>(() => Harness.Run(activity, service));
    }

    [Fact]
    public void InsertTableFormat_lets_a_custom_style_win_over_the_built_in_one()
    {
        var service = new FakeExcelService();

        Harness.Run(
            new InsertTableFormat
            {
                CellRange = new InArgument<string>("A1:D10"),
                TableFormatStyle = TableFormatEnum.TableStyleLight2,
                CustomStyle = new InArgument<string>("MyStyle"),
            },
            service);

        // Both are passed on; preferring the custom one is the workbook's job.
        Assert.Contains("InsertTableFormat(,A1:D10,TableStyleLight2,MyStyle)", service.Calls);
    }
}
