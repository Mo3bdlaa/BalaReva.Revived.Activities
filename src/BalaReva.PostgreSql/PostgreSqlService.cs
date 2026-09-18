using System.Data;
using Npgsql;

namespace BalaReva.PostgreSQL;

/// <summary>Runs commands through Npgsql.</summary>
/// <remarks>
/// Not covered by any automated test: no build agent has a PostgreSQL server. It is kept
/// to a mechanical translation for that reason, with the argument handling pushed up into
/// the activities, which are tested against a stand-in.
/// </remarks>
public sealed class PostgreSqlService : IPostgreSqlService
{
    /// <summary>The service the activities use when no extension is registered.</summary>
    public static IPostgreSqlService Instance { get; } = new PostgreSqlService();

    /// <inheritdoc />
    public DataTable ExecuteDataTable(PostgreSqlCommand command)
    {
        using var connection = Open(command);
        using var npgsql = Build(command, connection);
        using var reader = npgsql.ExecuteReader();

        var table = new DataTable("Result");
        table.Load(reader);
        return table;
    }

    /// <inheritdoc />
    public int ExecuteNonQuery(PostgreSqlCommand command)
    {
        using var connection = Open(command);
        using var npgsql = Build(command, connection);
        return npgsql.ExecuteNonQuery();
    }

    /// <inheritdoc />
    public object? ExecuteScalar(PostgreSqlCommand command)
    {
        using var connection = Open(command);
        using var npgsql = Build(command, connection);

        var value = npgsql.ExecuteScalar();
        return value is DBNull ? null : value;
    }

    private static NpgsqlConnection Open(PostgreSqlCommand command)
    {
        ArgumentNullException.ThrowIfNull(command);

        if (string.IsNullOrWhiteSpace(command.ConnectionString))
            throw new ArgumentException("ConnectionString is required.", nameof(command));

        var connection = new NpgsqlConnection(command.ConnectionString);
        connection.Open();
        return connection;
    }

    private static NpgsqlCommand Build(PostgreSqlCommand command, NpgsqlConnection connection)
    {
        if (string.IsNullOrWhiteSpace(command.CommandText))
            throw new ArgumentException("CmdText is required.", nameof(command));

        var npgsql = new NpgsqlCommand(command.CommandText, connection)
        {
            CommandType = command.CommandType,
        };

        foreach (var parameter in command.Parameters)
        {
            // A parameter object belongs to one command, and a workflow that reuses an
            // array across activities would otherwise hand over one Npgsql already owns.
            npgsql.Parameters.Add(parameter.Clone());
        }
        return npgsql;
    }
}
