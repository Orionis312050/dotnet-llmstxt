---
title: "Static Pages"
sidebar_position: "1"
---

# Static Pages

The `App` class lets you host web apps in the agent. This is useful for an efficient inner loop — you can build, deploy, and sideload both an agent and a Tab app inside Teams in a single step. It's also useful in production, since it makes it straightforward to host a simple experience such as an agent configuration page or a Dialog.

To host a static tab web app, call the SDK 2.0 `app.AddTab()` helper or SDK 2.1 ASP.NET route mapping with `app.MapGet(...)` function and provide an app name and a path to a folder containing an `index.html` file to be served up.

**SDK 2.0 (Legacy):**

```csharp
app.AddTab("myApp", "Web/bin");
```

**SDK 2.1 (current):**

```csharp
using Microsoft.AspNetCore.StaticFiles;

FileExtensionContentTypeProvider contentTypes = new();

app.MapGet("/tabs/myApp", () =>
{
    string index = Path.Combine(Directory.GetCurrentDirectory(), "Web", "bin", "index.html");
    contentTypes.TryGetContentType(index, out string? ct);
    return Results.File(File.OpenRead(index), ct ?? "text/html");
});

app.MapGet("/tabs/myApp/{*path}", (string? path) =>
{
    string root = Path.Combine(Directory.GetCurrentDirectory(), "Web", "bin");
    string full = Path.Combine(root, path ?? "index.html");
    contentTypes.TryGetContentType(full, out string? ct);
    return Results.File(File.OpenRead(full), ct ?? "text/html");
});
```

This registers a route that is hosted at `http://localhost:PORT/tabs/myApp` or `https://BOT_DOMAIN/tabs/myApp`.

## Additional resources

- For more details about Tab apps, see the [Tabs](../tabs.md) in-depth guide.
- For an example of hosting a Dialog, see the [Creating Dialogs](../dialogs/creating-dialogs.md) in-depth guide.
