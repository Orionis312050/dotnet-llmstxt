---
title: "Handling Tool Calls"
sidebar_position: "3"
---

# Handling Tool Calls

An interactive widget can call tools on your app.
When it does, the Teams client sends your app an `htmlwidget/calltool` invoke activity.
The SDK gives you a typed handler for these invokes and a typed response shape, so you never parse the raw invoke body.

:::warning Preview
HTML Widgets are in preview and the API surface may change. HTML Widgets are planned for public preview on 11/16.
:::

## The flow

1. Your app sends a widget whose HTML calls `tools/call` over the MCP Apps protocol.
2. The Teams client forwards that call to your app as an `htmlwidget/calltool` invoke.
3. Your handler inspects the tool name and arguments, does its work, and returns a result.
4. The Teams client delivers the result back to the widget, which re-renders.

## Registering a handler

Register a handler with `teams.OnWidgetCallTool`. The activity value is typed, so you get the tool `Name` and `Arguments` directly.

```csharp
using Microsoft.Teams.Apps.HtmlWidget;

teams.OnWidgetCallTool(async (context, cancellationToken) =>
{
    var request = context.Activity.Value;
    var toolName = request?.Name ?? "unknown";

    var response = toolName switch
    {
        "getTime" => new HtmlWidgetCallToolResponse
        {
            CallToolResult = new McpUiCallToolResult
            {
                Content = [new McpUiCallToolResultContent { Text = DateTime.UtcNow.ToString("HH:mm:ss") }],
                StructuredContent = new { time = DateTime.UtcNow.ToString("o") },
            }
        },
        _ => HtmlWidgetCallToolResponse.FromError($"Unknown tool: {toolName}"),
    };

    await Task.CompletedTask;
    return InvokeResponse.Ok(response);
});
```

## Returning a result

A tool call result carries a list of content blocks, optional structured content the widget can render from, and an `isError` flag.
Teams currently renders `text` content; the other MCP content types (`image`, `audio`, `resource`) are defined for forward compatibility.

Return `InvokeResponse.Ok(...)` wrapping an `HtmlWidgetCallToolResponse` whose `ResponseType` is `htmlwidget/calltoolresult` and whose `CallToolResult` holds the payload.
The result's `Content` is a list of content blocks; `StructuredContent` holds data the widget can render from; `IsError` signals a failure.
`Content`, `StructuredContent` and `IsError` are always serialized, defaulting to an empty array, an empty object and `false`, because an omitted field can reach the widget as `null`.
For the common text-only case, use the `FromText` and `FromError` factory methods.
.NET has no `buildCallToolResult` equivalent, because `McpUiCallToolResult` applies those defaults itself: `Content` and `StructuredContent` fall back to an empty array and object when you assign `null`, and `IsError` is a `bool` that defaults to `false`.

```csharp
var response = new HtmlWidgetCallToolResponse
{
    CallToolResult = new McpUiCallToolResult
    {
        Content = [new McpUiCallToolResultContent { Text = "Refreshed!" }],
        StructuredContent = new { counter = 1, lastAction = "refresh" },
    }
};

// Or, for a simple text result:
var quick = HtmlWidgetCallToolResponse.FromText("Refreshed!");
```

## Client-supplied call context

The Teams client adds an `__msteams__` entry to the `arguments` of every call, carrying the `mcpUrl` and `correlationId` from the widget payload when they were set. The client overwrites any value the widget supplied, and the SDK passes `arguments` to your handler untouched. Ignore `__msteams__` when you read your tool's arguments, validate the other arguments as untrusted input, never throw from the handler, and do not authorize from `activity.from`.

There is no separate .NET type for the entry. `Arguments` is typed `object?`, and it is a `JsonElement` when the widget sent an object.
Read the entry from `context.Activity.Value.Arguments` if you need `mcpUrl` or `correlationId`, and ignore it when you treat the remaining arguments as your tool's input:

```csharp
if (context.Activity.Value?.Arguments is JsonElement { ValueKind: JsonValueKind.Object } args
    && args.TryGetProperty("__msteams__", out var msteams)
    && msteams.TryGetProperty("correlationId", out var correlationId))
{
    // correlationId.GetString()
}
```

## Throttling in large contexts

In large or busy contexts the Teams platform may throttle widget tool calls, so your app can receive fewer `htmlwidget/calltool` invokes than the widget issued. This throttling happens on the platform before the invoke reaches your app, so there is no error or `429` status for your handler to catch and nothing for the SDK to retry. Design your tool handlers so that a dropped call is not fatal: keep them idempotent, and let the widget reflect its own state rather than assuming every call reaches your app.
