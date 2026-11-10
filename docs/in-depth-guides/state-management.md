---
title: "State Management"
sidebar_position: "10"
---

# State Management

:::info[SDK 2.1]
Built-in state management is a **Teams SDK 2.1** feature.
:::

The Teams SDK provides built-in, per-turn state for storing conversation and user data across activities. State is loaded before each activity handler runs, saved automatically after the turn, and exposed through `context.State`.

## Setup

State is disabled by default. Enable it with `UseState()`:

```csharp title="Program.cs"
using Microsoft.Teams.Apps;

WebApplicationBuilder builder = WebApplication.CreateSlimBuilder(args);
builder.Services.AddTeamsBotApplication(options =>
{
    options.UseState();
});

WebApplication app = builder.Build();
TeamsBotApplication teams = app.UseTeamsBotApplication();
```

Without a dedicated provider, state uses the in-memory `IDistributedCache` implementation. This is useful for local development, but values are lost when the process restarts and aren't shared across instances.

:::note
Registering an OAuth flow with `AddOAuthFlow()` automatically enables state so pending sign-ins can be associated with the correct flow.
:::

## Reading and writing state

Use `context.State.ConversationState` and `context.State.UserState` in an activity handler. Values must be JSON-serializable.

```csharp
teams.OnMessage(async (context, cancellationToken) =>
{
    // Write
    context.State.ConversationState.Set("lastMessage", context.Activity.Text ?? string.Empty);
    context.State.UserState.Set("messageCount",
        (context.State.UserState.Get<int>("messageCount")) + 1);

    // Read
    string? last = context.State.ConversationState.Get<string>("lastMessage");
    int count = context.State.UserState.Get<int>("messageCount");

    await context.SendAsync(
        $"Message #{count}. Last message was: {last}",
        cancellationToken);
});
```

Use `ContainsKey()` to check for a value, `Remove()` to remove one, and `Clear()` to remove all values. If you mutate an object or collection returned by `Get(string)`, call `Set()` with the updated value so the scope is marked for persistence.

Conversation state is shared by everyone in the current conversation. User state is scoped to the current sender **within that conversation** and is unavailable when an activity has no usable sender ID.

The SDK writes a scope to storage only when it changes during the current activity. After all handlers for that activity finish, the SDK saves those changes and closes the turn state. Read or update state only while handling the activity. Don't capture it for timers or background tasks because accessing it after the turn ends throws `InvalidOperationException`.

## Clearing state

Remove a value with `Remove()`, or clear a scope with `Clear()`. To remove both scopes from the backing store:

```csharp
await context.State.DeleteAsync(cancellationToken);
```

Values written after `DeleteAsync()` are saved normally at the end of the current turn.

## Scaling to distributed state

For production or multi-instance deployments, register a shared, durable `IDistributedCache` provider, such as Redis, SQL Server, or Azure Cache for Redis, before calling `UseState()`. The state API used by handlers doesn't change:

```csharp title="Program.cs"
using Microsoft.Teams.Apps;
using StackExchange.Redis;

WebApplicationBuilder builder = WebApplication.CreateSlimBuilder(args);

// Register Redis — UseState() picks this up automatically
builder.Services.AddStackExchangeRedisCache(options =>
{
    options.Configuration = builder.Configuration["Redis:ConnectionString"];
});

builder.Services.AddTeamsBotApplication(options =>
{
    options.UseState();
});
```

The SDK serializes each scope as UTF-8 JSON bytes and replaces the complete scope on save, so concurrent turns use last-writer-wins semantics. Configure cache entry expiration and the key prefix through `UseState()`. Configure any provider-specific options when registering the `IDistributedCache` implementation.
