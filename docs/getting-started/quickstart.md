---
title: "Quickstart"
sidebar_position: "1"
---

# Quickstart

Get started with Teams SDK quickly using the Teams Developer CLI.

## Set up a new project

### Prerequisites

- **.NET** v8 or higher. Install or upgrade from [dotnet.microsoft.com](https://dotnet.microsoft.com/en-us/download).

## Instructions

### Install the Teams Developer CLI

Install `teams` globally:

```sh
npm install -g @microsoft/teams.cli
teams --version
```

:::info
The [Teams Developer CLI](../developer-tools/teams-developer-cli.md) is the command-line tool for scaffolding, registering, and managing Teams apps.
:::

## Creating Your First Agent

Let's begin by creating a simple echo agent that responds to messages. Run:

```sh
teams project new csharp quote-agent --template echo
```

This command:

1. Creates a new directory called `QuoteAgent`.
2. Bootstraps the echo agent template files into your project directory.

> The `echo` template creates a basic agent that repeats back any message it receives - perfect for learning the fundamentals.

## Running your agent

1. Navigate to your new agent's directory:

```sh
cd QuoteAgent/QuoteAgent
```

2. Install the dependencies:

```sh
dotnet restore
```

3. Start the development server:

```sh
dotnet run
```

1. In the console, you should see a similar output:

```sh
info: Microsoft.Hosting.Lifetime[14]
      Now listening on: http://localhost:3978
info: Microsoft.Hosting.Lifetime[0]
      Application started. Press Ctrl+C to shut down.
```

The HTTP server is now listening on port `3978`. To test your agent locally without sideloading it into Teams, use the **[Microsoft 365 Agents Playground](../developer-tools/microsoft-365-agents-playground.md)**.

**SDK 2.0 (Legacy):**

The playground sends unauthenticated requests, which a default `builder.AddTeams()` rejects when no credentials are configured. For local testing, enable `skipAuth` so your agent accepts them:

```csharp title="Program.cs"
builder.AddTeams(skipAuth: true);
```

:::warning
Only use `skipAuth` for local development — never in production, as it disables inbound request authentication.
:::

**SDK 2.1 (current):**

The playground sends unauthenticated requests, and SDK 2.1 rejects them by default. For local testing, enable `DangerouslyAllowUnauthenticatedRequests` in your launch settings:

```json title="Properties/launchSettings.json"
{
  "profiles": {
    "https": {
      "environmentVariables": {
        "AzureAd__DangerouslyAllowUnauthenticatedRequests": "true"
      }
    }
  }
}
```

:::warning
Only use `DangerouslyAllowUnauthenticatedRequests` for local development — never in production, as it disables inbound request authentication.
:::

Install the playground globally:

```sh
npm install -g @microsoft/m365agentsplayground
```

Then, with your agent still running, open a second terminal and launch the playground pointed at your agent:

```sh
agentsplayground -e http://localhost:3978/api/messages -c emulator
```

The playground opens at [http://localhost:56150](http://localhost:56150). Send a message in the compose box and your agent's reply renders inline.

![Microsoft 365 Agents Playground showing a user message 'hello!' and an agent reply 'you said "hello!"'.](/screenshots/agents-playground-echo-chat.png)

## Add to an Existing Project

If you already have a project and want to add Teams support, install the SDK directly:

```sh
dotnet add package Microsoft.Teams.Apps
```
// Then register the Teams services on your existing app.

```csharp
using Microsoft.Teams.Apps;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddTeamsBotApplication();
var app = builder.Build();

// Maps POST /api/messages onto your existing ASP.NET Core app
TeamsBotApplication teams = app.UseTeamsBotApplication();

teams.OnMessage(async (context, cancellationToken) =>
{
    await context.SendAsync($"you said: {context.Activity.Text}", cancellationToken);
});

app.Run();
```
`app.UseTeamsBotApplication()` registers the Teams endpoint onto your existing ASP.NET Core app.

## Next steps

After creating and running your first agent, read about [the code basics](./code-basics.md) to better understand its components and structure.

Otherwise, if you want to run your agent in Teams, you can check out the [Running in Teams](running-in-teams) guide.

## Resources

- [Teams Developer CLI documentation](../developer-tools/teams-developer-cli.md)
- [Microsoft 365 Agents Playground](../developer-tools/microsoft-365-agents-playground.md)
- [Teams manifest schema](https://learn.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema)
- [Teams sideloading](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/apps-upload)
