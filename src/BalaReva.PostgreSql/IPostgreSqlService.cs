using System.Data;
using Npgsql;

namespace BalaReva.PostgreSQL;

/// <summary>Runs commands against a PostgreSQL database.</summary>
/// <remarks>
/// The activities talk to this rather than to Npgsql directly, so argument handling and
/// output mapping can be tested without a server. The real implementation is a thin
/// translation, as with the Office packages, and for the same reason: no build agent has
/// a database.
/// </remarks>
public interface IPostgreSqlService
{
    /// <summary>Runs a command and returns its rows.</summary>
    DataTable ExecuteDataTable(PostgreSqlCommand command);

    /// <summary>Runs a command and returns how many rows it affected.</summary>
    int ExecuteNonQuery(PostgreSqlCommand command);

    /// <summary>Runs a command and returns the first column of its first row.</summary>
    object? ExecuteScalar(PostgreSqlCommand command);
}

/// <summary>What the activities ask the service to run.</summary>
/// <remarks>Not part of the published surface; it keeps the service signatures readable.</remarks>
public sealed class PostgreSqlCommand
{
    /// <summary>Connection string for the database.</summary>
    public string ConnectionString { get; set; } = string.Empty;

    /// <summary>SQL text, or the name of a stored procedure.</summary>
    public string CommandText { get; set; } = string.Empty;

    /// <summary>Whether CommandText is SQL or the name of a stored procedure.</summary>
    public CommandType CommandType { get; set; } = CommandType.Text;

    /// <summary>
    /// Parameters to bind. Passing values here rather than building them into the SQL is
    /// what keeps a query free of injection.
    /// </summary>
    public NpgsqlParameter[] Parameters { get; set; } = [];
}
