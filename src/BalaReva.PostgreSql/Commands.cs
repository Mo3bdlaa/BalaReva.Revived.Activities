using System.Activities;
using System.ComponentModel;
using System.Data;

namespace BalaReva.PostgreSQL;

/// <summary>Runs a query and returns its rows.</summary>
[DisplayName("Execute DataTable")]
[Description("Runs a query and returns its rows as a DataTable.")]
public sealed class ExecuteDataTable : BaseData
{
    /// <summary>The rows the query returned.</summary>
    [Category("Output")]
    [DisplayName("Output DataTable")]
    [Description("The rows the query returned.")]
    public OutArgument<DataTable> OutputDataTable { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IPostgreSqlService service, PostgreSqlCommand command) =>
        OutputDataTable.Set(context, service.ExecuteDataTable(command));
}

/// <summary>Runs a command and returns how many rows it affected.</summary>
[DisplayName("Execute Non Query")]
[Description("Runs a command and returns how many rows it affected.")]
public sealed class ExecuteNonQuery : BaseData
{
    /// <summary>How many rows the command affected.</summary>
    [Category("Output")]
    [DisplayName("Result Value")]
    [Description("How many rows the command affected.")]
    public OutArgument<int> ResultValue { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IPostgreSqlService service, PostgreSqlCommand command) =>
        ResultValue.Set(context, service.ExecuteNonQuery(command));
}

/// <summary>Runs a command and returns the first column of its first row.</summary>
[DisplayName("Execute Scalar")]
[Description("Runs a command and returns the first column of its first row.")]
public sealed class ExecuteScalar : BaseData
{
    /// <summary>The single value the command returned, or null.</summary>
    [Category("Output")]
    [DisplayName("Result Object")]
    [Description("The single value the command returned, or null when it returned none.")]
    public OutArgument<object> ResultObject { get; set; } = null!;

    /// <inheritdoc />
    protected override void ExecuteWork(
        CodeActivityContext context, IPostgreSqlService service, PostgreSqlCommand command) =>
        ResultObject.Set(context, service.ExecuteScalar(command)!);
}
