using System.Text.Json;
using AwesomeAssertions;
using Microsoft.Extensions.Logging.Abstractions;
using ModelContextProtocol.Protocol;
using NSubstitute;

namespace MicrosoftMcp.Todo.Tests;

internal static class ToolResults
{
    private static readonly JsonSerializerOptions Read =
        new(ToolResult.Json) { PropertyNameCaseInsensitive = true };

    internal static T Ok<T>(CallToolResult result)
    {
        Assert.False(result.IsError is true);
        return JsonSerializer.Deserialize<T>(ToolResult.ReadText(result), Read)
            ?? throw new InvalidOperationException("Tool result was null JSON.");
    }

    internal static string Fail(CallToolResult result)
    {
        Assert.True(result.IsError is true);
        return ToolResult.ReadText(result);
    }
}

public sealed class TodoToolsTests
{
    private static TodoTools Create(IGraphTodoService todo) =>
        new(todo, NullLogger<TodoTools>.Instance);

    [Fact]
    public async Task Harvest_commitments_with_fake()
    {
        var tools = Create(new FakeGraphTodoService());

        var lists = ToolResults.Ok<List<TodoListInfo>>(await tools.todo_list_lists());
        lists.Should().ContainSingle(l => l.Id == "l-tasks");
        lists.Should().ContainSingle(l => l.IsShared == true);

        // Default list hides completed tasks.
        var tasks = ToolResults.Ok<List<TodoTaskSummary>>(await tools.todo_list_tasks());
        tasks.Should().ContainSingle(t => t.Id == "t-1");
        tasks.Should().NotContain(t => t.Id == "t-done");

        var all = ToolResults.Ok<List<TodoTaskSummary>>(
            await tools.todo_list_tasks(includeCompleted: true));
        all.Should().Contain(t => t.Id == "t-done");

        var added = ToolResults.Ok<TodoTaskDetail>(
            await tools.todo_add_task("Angebot bis Freitag senden", dueDateTime: "2026-10-15T17:00:00Z", importance: "high"));
        added.Title.Should().Be("Angebot bis Freitag senden");
        added.Status.Should().NotBe("Completed");

        var done = ToolResults.Ok<TodoTaskDetail>(
            await tools.todo_complete_task(added.Id));
        done.Status.Should().Be("Completed");
    }

    [Fact]
    public async Task Complete_with_explicit_list()
    {
        var tools = Create(new FakeGraphTodoService());

        var done = ToolResults.Ok<TodoTaskDetail>(
            await tools.todo_complete_task("t-shared", "l-shared"));
        done.Status.Should().Be("Completed");
        done.ListId.Should().Be("l-shared");
    }

    [Fact]
    public async Task Tool_errors_carry_codes_and_hints()
    {
        var tools = Create(new FakeGraphTodoService());

        ToolResults.Fail(await tools.todo_list_tasks("no-list")).Should().Contain("[todo-list-not-found]");
        ToolResults.Fail(await tools.todo_add_task("  ")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.todo_add_task("Hi", importance: "urgent")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.todo_add_task("Hi", dueDateTime: "kein-datum")).Should().Contain("[invalid-request]");
        ToolResults.Fail(await tools.todo_complete_task("nope")).Should().Contain("[todo-task-not-found]");
        ToolResults.Fail(await tools.todo_complete_task("t-1", "l-shared")).Should().Contain("[todo-task-not-found]");
    }

    [Fact]
    public async Task Unexpected_backend_failures_are_mapped_not_leaked()
    {
        IGraphTodoService failing = Substitute.For<IGraphTodoService>();
        failing.ListListsAsync(Arg.Any<CancellationToken>())
            .Returns<Task<IReadOnlyList<TodoListInfo>>>(_ => throw new HttpRequestException("no route"));
        var tools = Create(failing);

        var text = ToolResults.Fail(await tools.todo_list_lists());
        text.Should().Contain("[service-unavailable]");
        text.Should().Contain("Next:");
    }

    [Fact]
    public async Task Delegation_wiring_with_substitute()
    {
        IGraphTodoService todo = Substitute.For<IGraphTodoService>();
        todo.ListListsAsync(Arg.Any<CancellationToken>()).Returns(
            [new TodoListInfo("l1", "Tasks", true, false, "DefaultList")]);
        var tools = Create(todo);

        ToolResults.Ok<List<TodoListInfo>>(await tools.todo_list_lists())
            .Should().ContainSingle(l => l.Id == "l1");
    }
}
