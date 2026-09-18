# BalaReva Revived — PostgreSQL Activities

PostgreSQL activities for UiPath, reimplemented on .NET 8.

Type and property names match the published `BalaReva.PostgreSql.Activities` 2021.0.0, so
existing workflows keep binding. Note the casing: the namespace is `BalaReva.PostgreSQL`
while the assembly and package are `BalaReva.PostgreSql.Activities`. Both are as
published.

**This package targets plain `net8.0`.** Npgsql is managed and cross-platform, so these
activities work in a Studio *Cross-platform* project — something the original, which
shipped assemblies straight into `lib/` with no target framework at all, could not do.

## Why it was worth reimplementing

The published package pinned **Npgsql 4.0.7**, inside the affected range for
**CVE-2024-32655** (HIGH) — SQL injection through a protocol message size overflow, fixed
in 4.0.14. This package takes 10.0.3.

## The binding surface keeps an Npgsql type

`BaseData.Parameters` is an `InArgument<NpgsqlParameter[]>` in the published package, and
it stays one here: a workflow binds by type. Unlike the Office interop assemblies, Npgsql
is an ordinary maintained NuGet package that resolves cleanly on .NET 8, so there is no
reason to weaken it.

`BaseData` itself is a concrete `CodeActivity` rather than an abstract base, again as
published, so it appears in the toolbox in its own right and a workflow may have one on a
canvas. Running it executes the command and discards the result.

Parameters are cloned into each command before execution. An `NpgsqlParameter` belongs to
one command once added, so a workflow reusing an array across two activities would
otherwise hand over an object Npgsql already owns.

## What is not covered by tests

**No build agent has a PostgreSQL server**, so `PostgreSqlService` is not exercised by any
automated test. The activity layer above it is, against a recording stand-in: argument
validation, `ContinueOnError`, `Delay`, parameter passing and output mapping all sit on
that side of the boundary. The service itself is kept to a mechanical translation for
exactly that reason.
