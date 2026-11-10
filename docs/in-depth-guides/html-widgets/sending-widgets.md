---
title: "Sending Widgets"
sidebar_position: "2"
---

# Sending Widgets

The simplest widget is static HTML with no callbacks.
The SDK build helpers take a widget payload and produce a message with the correct `extendedmarkdown` text format and the fenced `html-widget` code block already assembled.

:::warning Preview
HTML Widgets are in preview and the API surface may change. HTML Widgets are planned for public preview on 11/16.
:::

:::warning Message export
There is a gap in the export APIs that prevents widget content from being exported. If your organization determines this is a problem, you can disable bots that may send widgets.
:::

## Building a widget message

Use `HtmlWidgetHelpers.BuildHtmlWidgetMessage` to turn a payload into a message you can send directly. The helper sets the text format to `extendedmarkdown` and wraps the payload in the fenced `html-widget` block. You only need to provide `Name`, `Html`, and `Domain`; the `Type` defaults to `widget/mcp-ui`.

```csharp
using Microsoft.Teams.Apps.HtmlWidget;

teams.OnMessage(async (context, cancellationToken) =>
{
    var message = HtmlWidgetHelpers.BuildHtmlWidgetMessage(
        new HtmlWidgetPayload
        {
            Name = "Simple Widget",
            Html = "<div>Hello from a widget</div>",
            Domain = "https://teams.microsoft.com",
        });

    await context.SendAsync(message, cancellationToken);
});
```

The examples on this page use minimal HTML to show the helpers. The SDK sends your HTML unchanged, so a widget that needs to report its size must also complete the handshake and report its height, as described in [Your widget owns the handshake and sizing](#your-widget-owns-the-handshake-and-sizing).

## Adding text before and after

You often want to introduce a widget with a sentence, or add a note after it.
The build helpers accept options for text to place before and after the widget block.

```csharp
var message = HtmlWidgetHelpers.BuildHtmlWidgetMessage(
    new HtmlWidgetPayload
    {
        Name = "Simple Widget",
        Html = "<div>Hello from a widget</div>",
        Domain = "https://teams.microsoft.com",
    },
    new HtmlWidgetMarkdownOptions
    {
        Before = "Here is a simple static widget:",
        After = "No callbacks needed for static content.",
    });

await context.SendAsync(message, cancellationToken);
```

## Building the markdown directly

The message helper covers the common case, but it produces one complete message from a single payload. Build just the markdown string when you need to do something the message helper cannot, such as combining more than one widget in a single message, embedding a widget inside text you are already assembling, or setting activity properties the helper does not touch (for example attachments or channel data).
Remember to set the text format to `extendedmarkdown` so the Teams client parses the widget block.

```csharp
using Microsoft.Teams.Apps.Schema;

// BuildHtmlWidgetMarkdown returns just the string, so you can compose several
// widgets into one message or splice a widget into text you already have.
var weather = HtmlWidgetHelpers.BuildHtmlWidgetMarkdown(
    new HtmlWidgetPayload
    {
        Name = "Weather",
        Html = "<div>Sunny, 72F</div>",
        Domain = "https://teams.microsoft.com",
    });

var forecast = HtmlWidgetHelpers.BuildHtmlWidgetMarkdown(
    new HtmlWidgetPayload
    {
        Name = "Forecast",
        Html = "<div>Rain tomorrow</div>",
        Domain = "https://teams.microsoft.com",
    });

var text = $"Today:\n\n{weather}\n\nLooking ahead:\n\n{forecast}";
await context.SendAsync(
    new MessageActivityInput { Text = text, TextFormat = TextFormats.ExtendedMarkdown },
    cancellationToken);
```

## Interactive widget HTML

The SDK sends your HTML exactly as authored and does not inject any script into it.
Every widget must carry its own `ui/initialize` handshake and `ui/notifications/size-changed` reporting, or it will not initialize or size correctly. This is described in [Your widget owns the handshake and sizing](#your-widget-owns-the-handshake-and-sizing).
