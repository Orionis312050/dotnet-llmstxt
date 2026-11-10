---
title: "Code Basics"
sidebar_position: "2"
---

# Code Basics

After following the guidance in [the quickstart](./quickstart.md) to create your first Teams application, let's review its structure and key components. This knowledge can help you build more complex applications as you progress.

## Project Structure

When you create a new Teams application, it generates a directory with this basic structure:

```
QuoteAgent/
├── Program.cs        # Main application startup code
├── appsettings.json    # Configuration (logging, auth, etc.)
```

- **Program.cs**: The entry point. A standard ASP.NET Core app that registers the Teams application into DI and wires it into the request pipeline.
- **appsettings.json**: Standard ASP.NET Core configuration — logging levels, bot credentials.

## Core Components

Let's break down the simple application from the [quickstart](./quickstart.md) into its core components.

### The App Class

The heart of an application is the `App` class. This class handles all incoming activities and manages the application's lifecycle. It also acts as a way to host your application service.

**SDK 2.0 (Legacy):**

```csharp title="Program.cs"
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Extensions;
using Microsoft.Teams.Plugins.AspNetCore.Extensions;

var builder = WebApplication.CreateBuilder(args);
builder.AddTeams();
var app = builder.Build();
var teams = app.UseTeams();

teams.OnMessage(async (context, cancellationToken) =>
{
    await context.Typing(cancellationToken: cancellationToken);
    await context.Send($"you said '{context.Activity.Text}'", cancellationToken);
});

app.Run();
```

The app configuration includes a variety of options that allow you to customize its behavior, including controlling the underlying server, authentication, and other settings.

**SDK 2.1 (current):**

```csharp title="Program.cs"
using Microsoft.Teams.Apps;

WebApplicationBuilder builder = WebApplication.CreateSlimBuilder(args);
builder.Services.AddTeamsBotApplication();
WebApplication app = builder.Build();

TeamsBotApplication teams = app.UseTeamsBotApplication();

teams.OnMessage(async (context, cancellationToken) =>
{
    await context.TypingAsync(cancellationToken);
    await context.SendAsync($"you said '{context.Activity.Text}'", cancellationToken);
});

app.Run();
```

`AddTeamsBotApplication()` registers everything the SDK needs into dependency injection. `UseTeamsBotApplication()` wires it into the ASP.NET Core pipeline and maps the `POST /api/messages` endpoint. From there, you register handlers for the activity types you care about.

### Plugins

:::note[Removed in SDK 2.1]
Plugins have been removed in SDK 2.1. Because your app is a standard ASP.NET Core app, extend it with regular **middleware** and **dependency injection** instead — for per-turn logic, implement `ITurnMiddleware` and register it with `teams.UseMiddleware(...)`.
:::

### Message Handling

Teams applications respond to various types of activities. The most basic is handling messages:

**SDK 2.0 (Legacy):**

```csharp title="Program.cs"
teams.OnMessage(async (context, cancellationToken) =>
{
    await context.TypingAsync(cancellationToken: cancellationToken);
    await context.SendAsync($"you said \"{context.Activity.Text}\"", cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp title="Program.cs"
teams.OnMessage(async (context, cancellationToken) =>
{
    await context.TypingAsync(cancellationToken);
    await context.SendAsync($"you said \"{context.Activity.Text}\"", cancellationToken);
});
```

You can also match specific messages with a regular expression — `teams.OnMessage("(?i)^help$", ...)` only fires for a case-insensitive `help`.

This code:

1. Listens for all incoming messages using the `OnMessage` handler.
2. Sends a typing indicator, which renders as an animated ellipsis (…) in the chat.
3. Responds by echoing back the received message.

:::info
Each activity type has a dedicated handler method (`OnMessage`, `OnMessageReaction`, `OnMembersAdded`, and so on) for type-safe, readable routing.
:::

### Application Lifecycle

Your application starts when you run:

**SDK 2.0 (Legacy):**

```csharp
WebApplication app = builder.Build();
app.UseTeamsBotApplication();
app.Run();
```

**SDK 2.1 (current):**

```csharp
WebApplication app = builder.Build();
app.UseTeamsBotApplication();
app.Run();
```

This code initializes your application server and, when configured for Teams, also authenticates it to be ready for sending and receiving messages.

## Next Steps

Now that you understand the basic structure of your Teams application, you're ready to [run it in Teams](running-in-teams). You'll use the Teams Developer CLI to register your bot and sideload it into Teams.

After that, you can:

- Add more activity handlers for different types of interactions. See [Listening to Activities](../essentials/listening-to-activities.md) for more details.
- Integrate with external services using the [API Client](../essentials/teams-api-client.md).
- Add interactive [cards](../in-depth-guides/adaptive-cards.md) and [dialogs](../in-depth-guides/dialogs.md).

Continue on to the next page to learn about these advanced features.

## Other Resources

- [Essentials](../essentials.md)
- [Teams concepts](../teams/user-authentication/sso-troubleshooting.md)
- [Teams developer tools](../developer-tools/exposing-your-local-app-with-a-tunnel.md)
