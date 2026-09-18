using System.Activities;
using System.ComponentModel;
using System.Data;
using Npgsql;

namespace BalaReva.PostgreSQL;

/// <summary>Shared arguments for every activity that runs a command.</summary>
/// <remarks>
/// Abstract, as the published package declared it: the three command activities inherit
/// these arguments, and a workflow binds them on the subclass rather than on this.
/// </remarks>
public abstract class BaseData : CodeActivity
{
    /// <summary>Connection string for the database.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Connection String")]
    [Description("Connection string for the PostgreSQL database.")]
    public InArgument<string> ConnectionString { get; set; } = null!;

    /// <summary>SQL text, or the name of a stored procedure.</summary>
    [RequiredArgument]
    [Category("Input")]
    [DisplayName("Command Text")]
    [Description("SQL to run, or the name of a stored procedure.")]
    public InArgument<string> CmdText { get; set; } = null!;

    /// <summary>Whether CmdText is SQL or the name of a stored procedure.</summary>
    [Category("Input")]
    [DisplayName("Command Type")]
    [Description("Whether Command Text is SQL or the name of a stored procedure.")]
    public CommandType CmdType { get; set; } = CommandType.Text;

    /// <summary>Parameters to bind to the command.</summary>
    [Category("Input")]
    [DisplayName("Parameters")]
    [Description("Parameters to bind. Use these rather than building values into the SQL.")]
    public InArgument<NpgsqlParameter[]> Parameters { get; set; } = null!;

    /// <summary>When true, failures are swallowed rather than faulting the workflow.</summary>
    [Category("Common")]
    [DisplayName("Continue On Error")]
    [Description("Swallow failures instead of faulting the workflow.")]
    public InArgument<bool> ContinueOnError { get; set; } = null!;

    /// <summary>Milliseconds to wait before running.</summary>
    [Category("Common")]
    [DisplayName("Delay")]
    [Description("Milliseconds to wait before this activity runs.")]
    public InArgument<short> Delay { get; set; } = null!;

    /// <inheritdoc />
    protected sealed override void Execute(CodeActivityContext context)
    {
        var delay = Delay?.Get(context) ?? 0;
        if (delay > 0) Thread.Sleep(TimeSpan.FromMilliseconds(delay));

        try
        {
            ExecuteWork(
                context,
                context.GetExtension<IPostgreSqlService>() ?? PostgreSqlService.Instance,
                Command(context));
        }
        catch (Exception) when (ContinueOnError?.Get(context) ?? false)
        {
            // Reported nowhere: this base has no ExecutionResult, which is how the
            // published package declared it.
        }
    }

    /// <summary>Runs the command and reports whatever the subclass reports.</summary>
    protected abstract void ExecuteWork(
        CodeActivityContext context, IPostgreSqlService service, PostgreSqlCommand command);

    /// <summary>Gathers the arguments into one command.</summary>
    protected PostgreSqlCommand Command(CodeActivityContext context)
    {
        var connectionString = ConnectionString?.Get(context);
        if (string.IsNullOrWhiteSpace(connectionString))
            throw new ArgumentException("Connection String is required.", nameof(ConnectionString));

        var commandText = CmdText?.Get(context);
        if (string.IsNullOrWhiteSpace(commandText))
            throw new ArgumentException("Command Text is required.", nameof(CmdText));

        return new PostgreSqlCommand
        {
            ConnectionString = connectionString,
            CommandText = commandText,
            CommandType = CmdType,
            Parameters = Parameters?.Get(context) ?? [],
        };
    }
}
