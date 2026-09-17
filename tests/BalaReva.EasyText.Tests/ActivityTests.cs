using System.Activities;
using System.Activities.Statements;
using BalaReva.EasyText.Base;
using BalaReva.EasyText.TextFile;
using BalaReva.EasyText.Utilities;

namespace BalaReva.EasyText.Tests;

/// <summary>
/// Every activity here runs through the real workflow runtime inside a real
/// TextScope, so the scope-to-child handle contract is covered too.
/// </summary>
public class ActivityTests
{
    private const string Sample = "alpha\nbravo\ncharlie\nbravo\n";

    private static BaseChildActivity Defaults(BaseChildActivity activity)
    {
        activity.ContinueOnError = new InArgument<bool>(false);
        activity.Delay = new InArgument<int>(0);
        return activity;
    }

    [Fact]
    public void ReadAll_returns_the_whole_file()
    {
        using var file = new ScopedFile(Sample);
        var activity = (ReadAll)Defaults(new ReadAll());

        var result = file.Run<string>(activity, (a, v) =>
        {
            ((ReadAll)a).Result = new OutArgument<string>(v);
            return a;
        });

        Assert.Equal(Sample, result);
    }

    [Fact]
    public void ReadAllArray_drops_the_trailing_newline()
    {
        using var file = new ScopedFile(Sample);
        var activity = (ReadAllArray)Defaults(new ReadAllArray());

        var result = file.Run<string[]>(activity, (a, v) =>
        {
            ((ReadAllArray)a).ArrayResult = new OutArgument<string[]>(v);
            return a;
        });

        Assert.Equal(["alpha", "bravo", "charlie", "bravo"], result);
    }

    [Fact]
    public void LineCount_counts_lines_not_newlines()
    {
        using var file = new ScopedFile(Sample);
        var activity = (LineCount)Defaults(new LineCount());

        var result = file.Run<int>(activity, (a, v) =>
        {
            ((LineCount)a).Result = new OutArgument<int>(v);
            return a;
        });

        Assert.Equal(4, result);
    }

    [Fact]
    public void FindLineIndex_returns_every_match_as_a_1_based_line_number()
    {
        using var file = new ScopedFile(Sample);
        var activity = (FindLineIndex)Defaults(new FindLineIndex());
        activity.Find = new InArgument<string>("bravo");

        var result = file.Run<int[]>(activity, (a, v) =>
        {
            ((FindLineIndex)a).Result = new OutArgument<int[]>(v);
            return a;
        });

        Assert.Equal([2, 4], result);
    }

    [Fact]
    public void FindLineIndex_honours_the_comparison_mode()
    {
        using var file = new ScopedFile("Alpha\nALPHA\nalpha\n");
        var activity = (FindLineIndex)Defaults(new FindLineIndex());
        activity.Find = new InArgument<string>("alpha");
        activity.TextComparison = StringComparisonEnum.OrdinalIgnoreCase;

        var result = file.Run<int[]>(activity, (a, v) =>
        {
            ((FindLineIndex)a).Result = new OutArgument<int[]>(v);
            return a;
        });

        Assert.Equal([1, 2, 3], result);
    }

    [Fact]
    public void FindText_returns_a_0_based_character_offset()
    {
        using var file = new ScopedFile(Sample);
        var activity = (FindText)Defaults(new FindText());
        activity.Find = new InArgument<string>("bravo");
        activity.StartIndex = new InArgument<int>(0);

        var result = file.Run<int>(activity, (a, v) =>
        {
            ((FindText)a).Result = new OutArgument<int>(v);
            return a;
        });

        Assert.Equal(Sample.IndexOf("bravo", StringComparison.Ordinal), result);
    }

    [Fact]
    public void FindText_returns_minus_one_when_there_is_no_match()
    {
        using var file = new ScopedFile(Sample);
        var activity = (FindText)Defaults(new FindText());
        activity.Find = new InArgument<string>("delta");
        activity.StartIndex = new InArgument<int>(0);

        var result = file.Run<int>(activity, (a, v) =>
        {
            ((FindText)a).Result = new OutArgument<int>(v);
            return a;
        });

        Assert.Equal(-1, result);
    }

    [Fact]
    public void ReadSpecificLine_reads_an_inclusive_1_based_range()
    {
        using var file = new ScopedFile(Sample);
        var activity = (ReadSpecificLine)Defaults(new ReadSpecificLine());
        activity.LineFrom = new InArgument<int>(2);
        activity.LineTo = new InArgument<int>(3);

        var result = file.Run<string[]>(activity, (a, v) =>
        {
            // Result is a required argument, so it has to be bound to something even
            // when this test only asserts on ArrayResult.
            var joined = new Variable<string>();
            ((ReadSpecificLine)a).Result = new OutArgument<string>(joined);
            ((ReadSpecificLine)a).ArrayResult = new OutArgument<string[]>(v);
            return new Sequence { Variables = { joined }, Activities = { a } };
        });

        Assert.Equal(["bravo", "charlie"], result);
    }

    [Fact]
    public void ReadSpecificLine_joins_the_range_with_the_file_line_ending()
    {
        using var file = new ScopedFile("alpha\r\nbravo\r\ncharlie\r\n");
        var activity = (ReadSpecificLine)Defaults(new ReadSpecificLine());
        activity.LineFrom = new InArgument<int>(1);
        activity.LineTo = new InArgument<int>(2);

        var result = file.Run<string>(activity, (a, v) =>
        {
            ((ReadSpecificLine)a).Result = new OutArgument<string>(v);
            return a;
        });

        Assert.Equal("alpha\r\nbravo", result);
    }

    [Fact]
    public void InsertLineAt_shifts_the_existing_lines_down()
    {
        using var file = new ScopedFile(Sample);
        var activity = (InsertLineAt)Defaults(new InsertLineAt());
        activity.LineAt = new InArgument<int>(2);
        activity.LineText = new InArgument<string>("inserted");

        file.Run(activity);

        Assert.Equal("alpha\ninserted\nbravo\ncharlie\nbravo\n", file.Contents);
    }

    [Fact]
    public void InsertLineAt_appends_at_one_past_the_end()
    {
        using var file = new ScopedFile(Sample);
        var activity = (InsertLineAt)Defaults(new InsertLineAt());
        activity.LineAt = new InArgument<int>(5);
        activity.LineText = new InArgument<string>("appended");

        file.Run(activity);

        Assert.Equal("alpha\nbravo\ncharlie\nbravo\nappended\n", file.Contents);
    }

    [Fact]
    public void DeleteAt_removes_the_addressed_line()
    {
        using var file = new ScopedFile(Sample);
        var activity = (DeleteAt)Defaults(new DeleteAt());
        activity.DeleteLineIndex = new InArgument<int>(2);

        file.Run(activity);

        Assert.Equal("alpha\ncharlie\nbravo\n", file.Contents);
    }

    [Fact]
    public void RemoveEmptyLine_drops_blank_and_whitespace_only_lines()
    {
        using var file = new ScopedFile("alpha\n\n   \nbravo\n");
        var activity = (RemoveEmptyLine)Defaults(new RemoveEmptyLine());

        file.Run(activity);

        Assert.Equal("alpha\nbravo\n", file.Contents);
    }

    [Fact]
    public void FindReplaceAllText_replaces_every_occurrence()
    {
        using var file = new ScopedFile(Sample);
        var activity = (FindReplaceAllText)Defaults(new FindReplaceAllText());
        activity.OldText = new InArgument<string>("bravo");
        activity.NewText = new InArgument<string>("BRAVO");

        file.Run(activity);

        Assert.Equal("alpha\nBRAVO\ncharlie\nBRAVO\n", file.Contents);
    }

    [Fact]
    public void FindReplaceAt_only_touches_the_requested_lines()
    {
        using var file = new ScopedFile(Sample);
        var activity = (FindReplaceAt)Defaults(new FindReplaceAt());
        activity.LineFrom = new InArgument<int>(1);
        activity.LineTo = new InArgument<int>(2);
        activity.OldText = new InArgument<string>("bravo");
        activity.NewText = new InArgument<string>("BRAVO");

        file.Run(activity);

        // The bravo on line 4 is outside the range and must survive.
        Assert.Equal("alpha\nBRAVO\ncharlie\nbravo\n", file.Contents);
    }
}
