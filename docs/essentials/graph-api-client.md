---
title: "Graph API Client"
sidebar_position: "8"
---

# Graph API Client

[Microsoft Graph](https://docs.microsoft.com/en-us/graph/overview) gives you access to the wider Microsoft 365 ecosystem. You can enrich your application with data from across Microsoft 365.

The SDK gives your application easy access to the Microsoft Graph API via the `Microsoft.Graph` package.

## Calling APIs

Microsoft Graph can be accessed by your application using its own application token, or by using the user's token. If you need access to resources that your application may not have, but your user does, you will need to use the user's scoped graph client. To grant explicit consent for your application to access resources on behalf of a user, follow the [auth guide](../in-depth-guides/user-authentication.md).

To access the graph using the Graph using the app, you may use the `GraphServiceClient` object to call the endpoint of your choice.

```csharp
using Azure.Identity;
using Microsoft.Graph;

var credential = new ClientSecretCredential(
    configuration["AzureAd:TenantId"],
    configuration["AzureAd:ClientId"],
    configuration["AzureAd:ClientCredentials:0:ClientSecret"]);

var graph = new GraphServiceClient(credential, ["https://graph.microsoft.com/.default"]);
var user = await graph.Me.GetAsync(cancellationToken: cancellationToken);
Console.WriteLine($"User ID: {user?.Id}");
Console.WriteLine($"User Display Name: {user?.DisplayName}");
Console.WriteLine($"User Email: {user?.Mail}");
Console.WriteLine($"User Job Title: {user?.JobTitle}");
```

To access Graph with the signed-in user's token, do this in a message handler:

**SDK 2.0 (Legacy):**

```csharp
app.OnMessage(async (context, cancellationToken) =>
{
    var user = await context.UserGraph.Me.GetAsync();
    Console.WriteLine($"User ID: {user.id}");
    Console.WriteLine($"User Display Name: {user.displayName}");
    Console.WriteLine($"User Email: {user.mail}");
    Console.WriteLine($"User Job Title: {user.jobTitle}");
});
```

Here, the "context.UserGraph" object is a scoped graph client for the user that sent the message.

**SDK 2.1 (current):**

```csharp
using System.Net.Http.Headers;

var flow = teams.GetOAuthFlow("graph");
teams.OnMessage(async (context, cancellationToken) =>
{
    var token = await flow.SignInAsync(context, cancellationToken);
    if (token is null) return;

    using var http = new HttpClient();
    http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

    var meJson = await http.GetStringAsync(
        "https://graph.microsoft.com/v1.0/me?$select=id,displayName,mail,jobTitle",
        cancellationToken);

    await context.SendAsync(meJson, cancellationToken);
});
```
