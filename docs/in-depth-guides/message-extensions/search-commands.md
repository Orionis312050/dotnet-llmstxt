---
title: "\ud83d\udd0d Search commands"
sidebar_position: "2"
---

# 🔍 Search commands

Message extension search commands allow users to search external systems and insert the results of that search into a message in the form of a card.

## Search command invocation locations

There are two different areas search commands can be invoked from:

1. Compose Area
2. Compose Box

### Compose Area and Box

![Screenshot of Teams with outlines around the 'Compose Box' (for typing messages) and the 'Compose Area' (the menu option next to the compose box that provides a search bar for actions and apps).](/screenshots/compose-area.png)

## Setting up your Teams app manifest

To use search commands you have to define them in the Teams app manifest. Here is an example:

```json
"composeExtensions": [
    {
        "botId": "${{BOT_ID}}",
        "commands": [
            {
                "id": "searchQuery",
                "context": [
                    "compose",
                    "commandBox"
                ],
                "description": "Test command to run query",
                "title": "Search query",
                "type": "query",
                "parameters": [
                    {
                        "name": "searchQuery",
                        "title": "Search Query",
                        "description": "Your search query",
                        "inputType": "text"
                    }
                ]
            }
        ]
    }
]
```

Here we are defining the `searchQuery` search (or query) command.

## Handle submission

Handle the search query submission when the `searchQuery` search command is invoked.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Api.Activities.Invokes.MessageExtensions;
using Microsoft.Teams.Api.MessageExtensions;
using Microsoft.Teams.Apps.Annotations;

//...

[MessageExtension.Query]
public Response OnMessageExtensionQuery(
    [Context] QueryActivity activity,
    [Context] IContext.Client client,
    [Context] ILogger log)
{
    log.Info("[MESSAGE_EXT_QUERY] Search query received");

    var commandId = activity.Value?.CommandId;
    var query = activity.Value?.Parameters?.FirstOrDefault(p => p.Name == "searchQuery")?.Value?.ToString() ?? "";

    log.Info($"[MESSAGE_EXT_QUERY] Command: {commandId}, Query: {query}");

    if (commandId == "searchQuery")
    {
        return CreateSearchResults(query, log);
    }

    return new Response
    {
        ComposeExtension = new Result
        {
            Type = ResultType.Result,
            AttachmentLayout = Layout.List,
            Attachments = new List()
        }
    };
}
```

**SDK 2.1 (current):**

```csharp
using Microsoft.Teams.Apps.MessageExtensions;

//...

bot.OnQuery(async (context, cancellationToken) =>
{
    MessageExtensionQuery? query = context.Activity.Value;
    string? commandId = query?.CommandId;
    string searchText = query?.Parameters?.FirstOrDefault(p => p.Name == "searchQuery")?.Value ?? "";

    if (commandId == "searchQuery")
    {
        return CreateSearchResults(searchText);
    }

    return MessageExtensionResponse.CreateBuilder()
        .WithType(MessageExtensionResponseTypes.Message)
        .WithText("Unknown command")
        .Build();
});
```

`CreateSearchResults()` method

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Api.MessageExtensions;
using Microsoft.Teams.Cards;
using Microsoft.Teams.Common;

//...

private static Response CreateSearchResults(string query, ILogger log)
{
    var attachments = new List();

    // Create simple search results
    for (int i = 1; i <= 5; i++)
    {
        var card = new AdaptiveCard
        {
            Body = new List
            {
                new TextBlock($"Search Result {i}")
                {
                    Weight = TextWeight.Bolder,
                    Size = TextSize.Large
                },
                new TextBlock($"Query: '{query}' - Result description for item {i}")
                {
                    Wrap = true,
                    IsSubtle = true
                }
            }
        };

        var previewCard = new ThumbnailCard()
        {
            Title = $"Result {i}",
            Text = $"This is a preview of result {i} for query '{query}'."
        };

        var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
        {
            ContentType = ContentType.AdaptiveCard,
            Content = card,
            Preview = new Microsoft.Teams.Api.MessageExtensions.Attachment
            {
                ContentType = ContentType.ThumbnailCard,
                Content = previewCard
            }
        };

        attachments.Add(attachment);
    }

    return new Response
    {
        ComposeExtension = new Result
        {
            Type = ResultType.Result,
            AttachmentLayout = Layout.List,
            Attachments = attachments
        }
    };
}
```

**SDK 2.1 (current):**

```csharp
using Microsoft.Teams.Apps.MessageExtensions;
using Microsoft.Teams.Apps.Schema;

//...

private static InvokeResponse CreateSearchResults(string query)
{
    var attachments = new List();

    for (int i = 1; i <= 5; i++)
    {
        // Thumbnail cards with a tap/invoke value to trigger OnSelectItem
        var previewCard = new
        {
            title = $"Result {i}",
            text = $"This is a preview of result {i} for query '{query}'.",
            tap = new
            {
                type = "invoke",
                value = new { itemIndex = i, query }
            }
        };

        attachments.Add(TeamsAttachment.CreateBuilder()
            .WithContent(previewCard)
            .WithContentType(AttachmentContentTypes.ThumbnailCard)
            .Build());
    }

    return MessageExtensionResponse.CreateBuilder()
        .WithType(MessageExtensionResponseTypes.Result)
        .WithAttachmentLayout(TeamsAttachmentLayouts.List)
        .WithAttachments([.. attachments])
        .Build();
}
```

To implement custom actions when a user clicks on a search result item, you can handle the select item event:

**SDK 2.0 (Legacy):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Api;
using Microsoft.Teams.Api.Activities.Invokes.MessageExtensions;
using Microsoft.Teams.Api.MessageExtensions;
using Microsoft.Teams.Apps.Annotations;
using Microsoft.Teams.Cards;

//...

[MessageExtension.SelectItem]
public Response OnMessageExtensionSelectItem(
    [Context] SelectItemActivity activity,
    [Context] IContext.Client client,
    [Context] ILogger log)
{
    log.Info("[MESSAGE_EXT_SELECT_ITEM] Item selection received");

    var selectedItem = activity.Value;
    log.Info($"[MESSAGE_EXT_SELECT_ITEM] Selected: {JsonSerializer.Serialize(selectedItem)}");

    return CreateItemSelectionResponse(selectedItem, log);
}

// Helper method to create item selection response
private static Response CreateItemSelectionResponse(object? selectedItem, ILogger log)
{
    var itemJson = JsonSerializer.Serialize(selectedItem);

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock("Item Selected")
            {
                Weight = TextWeight.Bolder,
                Size = TextSize.Large,
                Color = TextColor.Good
            },
            new TextBlock("You selected the following item:")
            {
                Wrap = true
            },
            new TextBlock(itemJson)
            {
                Wrap = true,
                FontType = FontType.Monospace,
                Separator = true
            }
        }
    };

    var attachment = new Microsoft.Teams.Api.MessageExtensions.Attachment
    {
        ContentType = new ContentType("application/vnd.microsoft.card.adaptive"),
        Content = card
    };

    return new Response
    {
        ComposeExtension = new Result
        {
            Type = ResultType.Result,
            AttachmentLayout = Layout.List,
            Attachments = new List { attachment }
        }
    };
}
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps.MessageExtensions;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Cards;

//...

bot.OnSelectItem(async (context, cancellationToken) =>
{
    JsonElement selectedItem = context.Activity.Value;
    string itemJson = JsonSerializer.Serialize(selectedItem);

    var card = new AdaptiveCard
    {
        Body = new List
        {
            new TextBlock("Item Selected")
            {
                Weight = TextWeight.Bolder,
                Size = TextSize.Large,
                Color = TextColor.Good
            },
            new TextBlock("You selected the following item:") { Wrap = true },
            new TextBlock(itemJson) { Wrap = true, FontType = FontType.Monospace, Separator = true }
        }
    };

    TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(JsonSerializer.SerializeToElement(card))
        .Build();

    return MessageExtensionResponse.CreateBuilder()
        .WithType(MessageExtensionResponseTypes.Result)
        .WithAttachmentLayout(TeamsAttachmentLayouts.List)
        .WithAttachments(attachment)
        .Build();
});
```

The search results include both a full adaptive card and a preview card. The preview card appears as a list item in the search command area:

![Screenshot of Teams showing a message extensions search menu open with list of search results displayed as preview cards.](/screenshots/preview-card.png)

When a user clicks on a list item the dummy adaptive card is added to the compose box:

![Screenshot of Teams showing the selected adaptive card added to the compose box.](/screenshots/card-in-compose.png)

To implement custom actions when a user clicks on a search result item, you can add the `tap` property to the preview card. This allows you to handle the click event with custom logic:

## Resources

- [Search command](https://learn.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/search-commands/define-search-command?tabs=Teams-toolkit%2Cdotnet)
- [Just-In-Time Install](https://learn.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/search-commands/universal-actions-for-search-based-message-extensions#just-in-time-install)
