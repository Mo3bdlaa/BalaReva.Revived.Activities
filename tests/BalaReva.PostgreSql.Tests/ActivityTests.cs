using System.Activities;
using System.Data;
using BalaReva.PostgreSQL;
using Npgsql;

namespace BalaReva.PostgreSql.Tests;

/// <summary>A stand-in database that records the commands it was given.</summary>
public sealed class FakePostgreSqlService : IPostgreSqlService
{
    public List<PostgreSqlCommand> Commands { get; } = [];

    public DataTable Table { get; set; } = new("Result");

    public int RowsAffected { get; set; } = 3;

    public object? Scalar { get; set; } = 42;

    public Exception? Throw { get; set; }

    private T Record<T>(PostgreSqlCommand command, T result)
    {
        Commands.Add(command);
        if (Throw is not null) throw Throw;
        return result;
    }

    public DataTable ExecuteDataTable(PostgreSqlCommand command) => Record(command, Table);

    public int ExecuteNonQuery(PostgreSqlCommand command) => Record(command, RowsAffected);

    public object? ExecuteScalar(PostgreSqlCommand command) => Record(command, Scalar);
}

/// <summary>
/// Argument handling and output mapping for the four activities. The service itself is
/// not covered: no build agent has a PostgreSQL server.
/// </summary>
public class ActivityTests
{
    [Fact]
    public void The_base_gathers_every_argument_into_one_command()
    {
        var service = new FakePostgreSqlService();
        var parameters = new[] { new NpgsqlParameter("id", 7) };

        Run(
            new ExecuteNonQuery
            {
                ConnectionString = new InArgument<string>("Host=db;Database=app"),
                CmdText = new InArgument<string>("delete from t where id = @id"),
                CmdType = CommandType.Text,
                Parameters = new InArgument<NpgsqlParameter[]>(_ => parameters),
                ResultValue = new OutArgument<int>(),
            },
            service);

        var command = Assert.Single(service.Commands);
        Assert.Equal("Host=db;Database=app", command.ConnectionString);
        Assert.Equal("delete from t where id = @id", command.CommandText);
        Assert.Equal(CommandType.Text, command.CommandType);
        Assert.Same(parameters, command.Parameters);
    }

    [Fact]
    public void A_stored_procedure_keeps_its_command_type()
    {
        var service = new FakePostgreSqlService();

        Run(
            new ExecuteNonQuery
            {
                ConnectionString = new InArgument<string>("Host=db"),
                CmdText = new InArgument<string>("refresh_totals"),
                CmdType = CommandType.StoredProcedure,
                ResultValue = new OutArgument<int>(),
            },
            service);

        Assert.Equal(CommandType.StoredProcedure, service.Commands[0].CommandType);
    }

    [Fact]
    public void No_parameters_arrives_as_an_empty_array_rather_than_null()
    {
        var service = new FakePostgreSqlService();

        Run(
            new ExecuteNonQuery
            {
                ConnectionString = new InArgument<string>("Host=db"),
                CmdText = new InArgument<string>("select 1"),
                ResultValue = new OutArgument<int>(),
            },
            service);

        Assert.Empty(service.Commands[0].Parameters);
    }

    [Fact]
    public void ExecuteDataTable_reports_the_rows()
    {
        var table = new DataTable("Customers");
        table.Columns.Add("Name", typeof(string));
        table.Rows.Add("Ada");

        var service = new FakePostgreSqlService { Table = table };
        var run = new Runner();

        var outputs = run.Invoke(
            new ExecuteDataTable
            {
                ConnectionString = new InArgument<string>("Host=db"),
                CmdText = new InArgument<string>("select * from customers"),
                OutputDataTable = run.Capture<DataTable>("OutputDataTable"),
            },
            service);

        Assert.Equal(
            "Customers",
            Assert.IsType<DataTable>(outputs.Values["OutputDataTable"]).TableName);
    }

    [Fact]
    public void ExecuteNonQuery_reports_the_row_count()
    {
        var service = new FakePostgreSqlService { RowsAffected = 12 };
        var run = new Runner();

        var outputs = run.Invoke(
            new ExecuteNonQuery
            {
                ConnectionString = new InArgument<string>("Host=db"),
                CmdText = new InArgument<string>("delete from t"),
                ResultValue = run.Capture<int>("ResultValue"),
            },
            service);

        Assert.Equal(12, outputs.Values["ResultValue"]);
    }

    [Fact]
    public void ExecuteScalar_reports_the_single_value()
    {
        var service = new FakePostgreSqlService { Scalar = "Ada" };
        var run = new Runner();

        var outputs = run.Invoke(
            new ExecuteScalar
            {
                ConnectionString = new InArgument<string>("Host=db"),
                CmdText = new InArgument<string>("select name from t limit 1"),
                ResultObject = run.Capture<object>("ResultObject"),
            },
            service);

        Assert.Equal("Ada", outputs.Values["ResultObject"]);
    }

    [Theory]
    [InlineData(null, "select 1")]
    [InlineData("", "select 1")]
    [InlineData("Host=db", null)]
    [InlineData("Host=db", "")]
    public void A_missing_connection_string_or_command_is_refused(string? connection, string? text)
    {
        var activity = new ExecuteNonQuery
        {
            ConnectionString = new InArgument<string>(connection!),
            CmdText = new InArgument<string>(text!),
            ResultValue = new OutArgument<int>(),
        };

        Assert.ThrowsAny<Exception>(() => Run(activity, new FakePostgreSqlService()));
    }

    [Fact]
    public void ContinueOnError_swallows_a_failure()
    {
        var service = new FakePostgreSqlService { Throw = new InvalidOperationException("no server") };

        var activity = new ExecuteNonQuery
        {
            ConnectionString = new InArgument<string>("Host=db"),
            CmdText = new InArgument<string>("select 1"),
            ContinueOnError = new InArgument<bool>(true),
            ResultValue = new OutArgument<int>(),
        };

        Run(activity, service);
        Assert.Single(service.Commands);
    }

    [Fact]
    public void Without_ContinueOnError_a_failure_faults_the_workflow()
    {
        var service = new FakePostgreSqlService { Throw = new InvalidOperationException("no server") };

        var activity = new ExecuteNonQuery
        {
            ConnectionString = new InArgument<string>("Host=db"),
            CmdText = new InArgument<string>("select 1"),
            ResultValue = new OutArgument<int>(),
        };

        Assert.ThrowsAny<Exception>(() => Run(activity, service));
    }

    [Fact]
    public void The_base_runs_on_its_own_because_the_published_package_let_it()
    {
        // BaseData is a concrete CodeActivity in the published package, so it appears in
        // the toolbox and a workflow may have one on a canvas. Running it executes the
        // command and discards the result.
        var service = new FakePostgreSqlService();

        Run(
            new BaseData
            {
                ConnectionString = new InArgument<string>("Host=db"),
                CmdText = new InArgument<string>("vacuum"),
            },
            service);

        Assert.Equal("vacuum", Assert.Single(service.Commands).CommandText);
    }

    private static void Run(Activity activity, IPostgreSqlService service)
    {
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

/// <summary>Copies a workflow variable into <see cref="Outputs"/> once the activity is done.</summary>
internal sealed class CaptureOutput<T>(Outputs sink, string name) : CodeActivity
{
    public InArgument<T> Value { get; set; } = null!;

    protected override void Execute(CodeActivityContext context) =>
        sink.Values[name] = context.GetValue(Value)!;
}

/// <summary>
/// Runs one activity and reads its output arguments back, which an activity cannot return
/// on its own.
/// </summary>
internal sealed class Runner
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

    public Outputs Invoke(Activity activity, IPostgreSqlService service)
    {
        var root = new System.Activities.Statements.Sequence { Activities = { activity } };
        foreach (var variable in _variables) root.Variables.Add(variable);
        foreach (var capture in _captures) root.Activities.Add(capture);

        var invoker = new WorkflowInvoker(root);
        invoker.Extensions.Add(service);
        invoker.Invoke();
        return _outputs;
    }
}
