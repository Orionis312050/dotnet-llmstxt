---
title: "\ud83d\udd12 User Authentication"
sidebar_position: "7"
---

# 🔒 User Authentication

At times agents must access secured online resources on behalf of the user, such as checking email, checking on flight status, or placing an order. To enable this, the user must authenticate their identity and grant consent for the application to access these resources. This process results in the application receiving a token, which the application can then use to access the permitted resources on the user's behalf.

:::info
This is an advanced guide. It is highly recommended that you are familiar with [Teams Core Concepts](../teams/teams-core-concepts.md) before attempting this guide.
:::

:::warning
User authentication does not work with the developer tools setup. You have to run the app in Teams. Follow [Quickstart: Register your app](../get-started/quickstart-register-your-app.md) to register and sideload your bot.
:::

:::info
It is possible to authenticate the user into [other auth providers](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-identity-providers?view=azure-bot-service-4.0&tabs=adv2%2Cga2#other-identity-providers) like Facebook, Github, Google, Dropbox, and so on.
:::

### Set up the OAuth connection

User authentication requires an **Azure-managed bot** (Teams-managed bots don't support OAuth connections). If you registered with `--teams-managed`, migrate first:

```sh
teams app bot migrate <appId> --subscription <id> --resource-group <your-resource-group>
```

Then follow the [User Authentication Setup guide](/cli/guides/user-authentication-setup) to configure the AAD app, create the Azure Bot OAuth connection, and update the manifest. The guide covers both SSO (silent token exchange) and generic OAuth.

:::tip
If you'd rather have an AI coding assistant run the setup, install the [`teams-dev` skill](../developer-tools/agent-skills.md) and ask it to "set up SSO for my Teams bot".
:::

## Configure the OAuth connection

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Plugins.AspNetCore.Extensions;

var builder = WebApplication.CreateBuilder(args);

var appBuilder = App.Builder()
    .AddOAuth("graph"); // default connection

builder.AddTeams(appBuilder);
var app = builder.Build();
var teams = app.UseTeams();
```

**SDK 2.1 (current):**

:::info
In SDK 2.1, you can register and use multiple OAuth connections in one bot (for example, Graph + GitHub) by adding multiple `AddOAuthFlow(...)` entries.
:::

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.OAuth;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddTeamsBotApplication(options =>
{
    options.AddOAuthFlow("graph", oauth =>
    {
        oauth.OAuthCardText = "Sign in to Microsoft Graph";
        oauth.SignInButtonText = "Sign in to Graph";
    });
    options.AddOAuthFlow("github", oauth =>
    {
        oauth.OAuthCardText = "Sign in to GitHub";
        oauth.SignInButtonText = "Sign in to GitHub";
    });
});

var app = builder.Build();
var teams = app.UseTeamsBotApplication();

OAuthFlow graphAuth = teams.GetOAuthFlow("graph");
OAuthFlow githubAuth = teams.GetOAuthFlow("github");
```

:::tip
Make sure you use the same name you used when creating the OAuth connection in the Azure Bot Service resource.
:::

:::note
In many templates, `graph` is the default name of the OAuth connection, but you can register the connection under any name — it just has to match the OAuth connection configured on the Azure Bot.
:::

## Signing In

Start sign-in from inside a route handler, for example when receiving the `/signin` message:

**SDK 2.0 (Legacy):**

:::note
This uses the Single Sign-On (SSO) authentication flow. To learn more about all the available flows and their differences see the [official documentation](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-authentication?view=azure-bot-service-4.0).
:::

```csharp
teams.OnMessage("/signin", async (context, cancellationToken) =>
{
    if (context.IsSignedIn)
    {
        await context.Send("you are already signed in!", cancellationToken);
        return;
    }

    await context.SignIn(new OAuthOptions
    {
        ConnectionName = "graph",
        OAuthCardText = "Sign in to your account",
        SignInButtonText = "Sign in"
    }, cancellationToken);
});
```

**SDK 2.1 (current):**

Call `SignInAsync` on the flow directly — each flow manages its own sign-in state independently.
  :::note
The graph connection uses the Single Sign-On (SSO) authentication flow whereas the Github connection uses the oauth flow. To learn more about all the available flows and their differences see the [official documentation](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-authentication?view=azure-bot-service-4.0).
:::

```csharp
teams.OnMessage("/signin graph", async (context, cancellationToken) =>
{
    string? token = await graphAuth.SignInAsync(context, cancellationToken);
    if (token is not null)
    {
        await context.SendAsync("you are already signed in to Graph!", cancellationToken);
    }
});

teams.OnMessage("/signin github", async (context, cancellationToken) =>
{
    string? token = await githubAuth.SignInAsync(context, cancellationToken);
    if (token is not null)
    {
        await context.SendAsync("you are already signed in to GitHub!", cancellationToken);
    }
});
```

## Handling sign-in completion

The sign-in method returns a token when the user already has one, and otherwise sends an OAuth card and returns nothing. The flow finishes on a later turn, so completion is delivered through a callback.

**SDK 2.0 (Legacy):**

```csharp
teams.OnSignIn(async (_, teamsEvent, cancellationToken) =>
{
    var context = teamsEvent.Context;
    await context.Send(
        $"Signed in using OAuth connection {context.ConnectionName}. Please type **/whoami** to see your profile or **/signout** to sign out.",
        cancellationToken);
});
```

**SDK 2.1 (current):**

Success and failure callbacks are scoped per flow — no branching on connection name needed.
```csharp
graphAuth.OnSignInComplete(async (context, tokenResponse, cancellationToken) =>
{
    await context.SendAsync(
        $"Signed in to Graph ({tokenResponse.ConnectionName}). Type **/whoami** to continue.",
        cancellationToken);
});

githubAuth.OnSignInComplete(async (context, tokenResponse, cancellationToken) =>
{
    await context.SendAsync(
        $"Signed in to GitHub ({tokenResponse.ConnectionName}).",
        cancellationToken);
});
```

## Start using the client

:::note
The default graph configuration requests the `User.ReadBasic.All` permission. It is possible to request other permissions by modifying the App Registration for the bot on Azure.
:::

**SDK 2.0 (Legacy):**

```csharp
teams.OnMessage("/whoami", async (context, cancellationToken) =>
{
    if (!context.IsSignedIn)
    {
        await context.Send("you are not signed in. Please type **/signin** to sign in.", cancellationToken);
        return;
    }

    var me = await context.GetUserGraphClient().Me.GetAsync(cancellationToken: cancellationToken);
    await context.Send($"user \"{me!.DisplayName}\" signed in.", cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
using System.Net.Http.Headers;

teams.OnMessage("/whoami", async (context, cancellationToken) =>
{
    string? token = await graphAuth.SignInAsync(context, cancellationToken);
    if (token is null) return; // OAuth card sent / waiting for callback turn

    using var http = new HttpClient();
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

    string meJson = await http.GetStringAsync(
        "https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName",
        cancellationToken);

    await context.SendAsync(meJson, cancellationToken);
});

teams.OnMessage("(?i)^my gh user$", async (context, cancellationToken) =>
{
    string? token = await githubAuth.SignInAsync(context, cancellationToken);
    if (token is null) return;

    using HttpClient http = new();
    http.DefaultRequestHeaders.Authorization = new("Bearer", token);
    http.DefaultRequestHeaders.UserAgent.ParseAdd("TeamsBot/1.0");

    string response = await http.GetStringAsync(
        "https://api.github.com/user", ct);
    await context.SendAsync($"Your GitHub user :\n```json\nresponse\n```", ct);
});

```

## Signing Out

You can signout by calling the `signout` method, this will remove the token from the User Token service cache

**SDK 2.0 (Legacy):**

```csharp
teams.OnMessage("/signout", async (context, cancellationToken) =>
{
    if (!context.IsSignedIn)
    {
        await context.Send("you are not signed in!", cancellationToken);
        return;
    }

    await context.SignOut(cancellationToken: cancellationToken);
    await context.Send("you have been signed out!", cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
teams.OnMessage("/signout graph", async (context, cancellationToken) =>
{
    await graphAuth.SignOutAsync(context, cancellationToken);
    await context.SendAsync("you have been signed out from Graph!", cancellationToken);
});

teams.OnMessage("/signout github", async (context, cancellationToken) =>
{
    await githubAuth.SignOutAsync(context, cancellationToken);
    await context.SendAsync("you have been signed out from GitHub!", cancellationToken);
});
```

## Using multiple OAuth connections

An agent may need more than one identity — for example Graph for the user's profile and GitHub for their repos. Each connection is registered separately, so tokens, callbacks and sign-out are scoped to the connection that owns them.

```csharp
builder.Services.AddTeamsBotApplication(options =>
{
    options.AddOAuthFlow("graph", oauth => oauth.OAuthCardText = "Sign in with Microsoft");
    options.AddOAuthFlow("github", oauth => oauth.OAuthCardText = "Sign in with GitHub");
});

OAuthFlow graphAuth = teams.GetOAuthFlow("graph");
OAuthFlow githubAuth = teams.GetOAuthFlow("github");

teams.OnMessage("/graph", async (context, cancellationToken) =>
{
    string? token = await graphAuth.SignInAsync(context, cancellationToken);
    if (token is not null)
    {
        await SendGraphProfileAsync(context, token, cancellationToken);
    }
});

teams.OnMessage("/github", async (context, cancellationToken) =>
{
    string? token = await githubAuth.SignInAsync(context, cancellationToken);
    if (token is not null)
    {
        await SendGitHubProfileAsync(context, token, cancellationToken);
    }
});

graphAuth.OnSignInComplete(async (context, tokenResponse, cancellationToken) =>
    await SendGraphProfileAsync(context, tokenResponse.Token, cancellationToken));

githubAuth.OnSignInComplete(async (context, tokenResponse, cancellationToken) =>
    await SendGitHubProfileAsync(context, tokenResponse.Token, cancellationToken));

teams.OnMessage("/signout github", async (context, cancellationToken) =>
{
    await githubAuth.SignOutAsync(context, cancellationToken);
    await context.SendAsync("Signed out of GitHub.", cancellationToken);
});
```

Signing out of one connection leaves the other signed in.

:::info
Each connection must exist on the Azure Bot resource under the name you register it with. Follow the [User Authentication Setup guide](/cli/guides/user-authentication-setup) once per connection.
:::

## Checking connection status

One call returns the status of every registered connection, so you don't have to probe them individually.

```csharp
using Microsoft.Teams.Core;

teams.OnMessage("/status", async (context, cancellationToken) =>
{
    IList statuses = await graphAuth.GetConnectionStatusAsync(context, cancellationToken);
    IEnumerable<string> lines = statuses.Select(status =>
        $"- `{status.ConnectionName}`: {(status.HasToken == true ? "signed in" : "signed out")}");
    await context.SendAsync(string.Join("\n", lines), cancellationToken);
});
```

`GetConnectionStatusAsync` reports every connection registered on the bot, so calling it on any flow returns the same list.

## Resuming Pending Messages After Sign-In

When a user isn't signed in and your message handler calls the sign-in method, an OAuth card is sent and the current turn ends. The sign-in completes on a *separate turn* — meaning the original message text is not available in the sign-in success context.

To avoid ignoring what the user originally asked, store the pending message before initiating sign-in, then retrieve and process it once sign-in succeeds:

**SDK 2.0 (Legacy):**

```csharp
using System.Collections.Concurrent;

var pendingMessages = new ConcurrentDictionary<string, string>();

teams.OnMessage(async (context, cancellationToken) =>
{
    if (!context.IsSignedIn)
    {
        var userId = context.Activity.From?.Id ?? string.Empty;
        pendingMessages[userId] = context.Activity.Text ?? string.Empty;

        await context.SignIn(cancellationToken: cancellationToken);
        return;
    }

    await ProcessMessage(context.Activity.Text, context, cancellationToken);
});

teams.OnSignIn(async (_, teamsEvent, cancellationToken) =>
{
    var context = teamsEvent.Context;
    var userId = context.Activity.From?.Id ?? string.Empty;

    if (pendingMessages.TryRemove(userId, out var text))
    {
        await context.Send("Successfully signed in! Processing your original request...", cancellationToken);
        await ProcessMessage(text, context, cancellationToken);
    }
});
```

**SDK 2.1 (current):**

```csharp
using System.Collections.Concurrent;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.OAuth;

var pendingMessages = new ConcurrentDictionary<string, string>();

teams.OnMessage("(?i)^my ad user$", async (context, cancellationToken) =>
{
    var userId = context.Activity.From?.Id ?? string.Empty;
    string? token = await graphAuth.SignInAsync(context, cancellationToken);
    if (token is null)
    {
        pendingMessages[userId] = context.Activity.Text ?? string.Empty;
        return;
    }

    await ProcessGraphMessage(context.Activity.Text, context, token, cancellationToken);
});

graphAuth.OnSignInComplete(async (context, tokenResponse, cancellationToken) =>
{
    var userId = context.Activity.From?.Id ?? string.Empty;
    if (pendingMessages.TryRemove(userId, out var text))
    {
        await context.SendAsync("Successfully signed in to Graph! Processing your request...", cancellationToken);
        await ProcessGraphMessage(text, context, tokenResponse.Token, cancellationToken);
    }
});
```

:::tip
For production apps, consider using a persistent store (database, Redis, etc.) instead of an in-memory map so pending messages survive restarts. You should also implement expiration or cleanup logic (e.g., a TTL) to discard stale entries when sign-in is cancelled, times out, or fails.
:::

## Handling Sign-In Failures

When using SSO, if the token exchange fails Teams sends a `signin/failure` invoke activity to your app. The SDK includes a built-in default handler that logs a warning with actionable troubleshooting guidance. You can optionally register your own handler to customize the behavior:

**SDK 2.0 (Legacy):**

```csharp
teams.OnSignInFailure(async (context, cancellationToken) =>
{
    var failure = context.Activity.Value;
    context.Log.Error($"sign-in failed: {failure?.Code} - {failure?.Message}");
    await context.Send("Sign-in failed.", cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
graphAuth.OnSignInFailure(async (context, failure, cancellationToken) =>
{
    await context.SendAsync(
        $"Graph sign-in failed: {failure?.Code} - {failure?.Message}",
        cancellationToken);
});

githubAuth.OnSignInFailure(async (context, failure, cancellationToken) =>
{
    await context.SendAsync(
        $"GitHub sign-in failed: {failure?.Code} - {failure?.Message}",
        cancellationToken);
});
```

:::tip
The most common failure codes are `installedappnotfound` (bot app not installed for the user) and `resourcematchfailed` (Token Exchange URL doesn't match the Application ID URI). See [SSO Setup - Troubleshooting](../teams/user-authentication/sso-setup.md#troubleshooting) for a full list of failure codes and troubleshooting steps.
:::

:::warning
A `signin/failure` activity doesn't name its connection. The SDK attributes it from the pending sign-in recorded when the flow started; if that can't be resolved, *every* registered connection's failure handler is notified. Check the connection first if your handler does something destructive.

Registering a flow also turns on per-turn state unless you configure `state` yourself, defaulting to local, in-process storage. Configure shared state storage before running more than one instance — otherwise pending sign-in attribution and token-exchange dedup don't cross processes, and sign-in fails intermittently.
:::

## Resources

[User Authentication Basics](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-authentication?view=azure-bot-service-4.0)
