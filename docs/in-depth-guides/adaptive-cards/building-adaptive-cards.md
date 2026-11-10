---
title: "Building Adaptive Cards"
sidebar_position: "1"
---

# Building Adaptive Cards

Adaptive Cards are JSON payloads that describe rich, interactive UI fragments.

With `Microsoft.Teams.Cards` you can build these cards entirely in C# while enjoying full IntelliSense and compiler safety.

## The Builder Pattern

`Microsoft.Teams.Cards` exposes small **builder helpers** including `AdaptiveCard`, `TextBlock`, `ToggleInput`, `ExecuteAction`, _etc._

Each helper wraps raw JSON and provides fluent, chainable methods that keep your code concise and readable.

```csharp
using Microsoft.Teams.Cards;

var card = new AdaptiveCard
{
    Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
    Body = new List
    {
        new TextBlock("Hello world")
        {
            Wrap = true,
            Weight = TextWeight.Bolder
        },
        new ToggleInput("Notify me")
        {
            Id = "notify"
        }
    },
    Actions = new List
    {
        new ExecuteAction
        {
            Title = "Submit",
            Data = new Union<string, SubmitActionData>(new SubmitActionData
            {
                NonSchemaProperties = new Dictionary<string, object?>
                {
                    { "action", "submit_basic" }
                }
            }),
            AssociatedInputs = AssociatedInputs.Auto
        }
    }
};
```

Benefits:

| Benefit     | Description                                                                   |
| ----------- | ----------------------------------------------------------------------------- |
| Readability | No deep JSON trees—just chain simple methods.                                 |
| Re‑use      | Extract snippets to functions or classes and share across cards.              |
| Safety      | Builders validate every property against the Adaptive Card schema (see next). |

:::info
The builder helpers use strongly-typed interfaces. Use IntelliSense (Ctrl+Space) or "Go to Definition" (F12) in your IDE to explore available types and properties. Source code lives in the `Microsoft.Teams.Cards` namespace.
:::

## Type‑safe Authoring & IntelliSense

The package bundles the **Adaptive Card v1.5 schema** as strict C# types.
While coding you get:

- **Autocomplete** for every element and attribute.
- **In‑editor validation**—invalid enum values or missing required properties produce build errors.
- Automatic upgrades when the schema evolves; simply update the package.

```csharp
// "Huge" is not a valid size for TextBlock - this will cause a compilation error
var textBlock = new TextBlock("Test")
{
    Wrap = true,
    Weight = TextWeight.Bolder,
    Size = "Huge" // This is invalid - should be TextSize enum
};
```

## The Visual Designer

Prefer a drag‑and‑drop approach? Use [Microsoft's Adaptive Card Designer](https://adaptivecards.microsoft.com/designer.html):

1. Add elements visually until the card looks right.
2. Copy the JSON payload from the editor pane.
3. Paste the JSON into your project **or** convert it to builder calls:

**SDK 2.0 (Legacy):**

```csharp
    var cardJson = """
    {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "verticalContentAlignment": "center",
                        "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://aka.ms/AAp9xo4",
                                "size": "Small",
                                "altText": "Portrait of David Claux"
                            }
                        ],
                        "width": "auto"
                    },
                    {
                        "type": "Column",
                        "spacing": "medium",
                        "verticalContentAlignment": "center",
                        "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "David Claux",
                                "wrap": true
                            }
                        ],
                        "width": "auto"
                    },
                    {
                        "type": "Column",
                        "spacing": "medium",
                        "verticalContentAlignment": "center",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Principal Platform Architect at Microsoft",
                                "isSubtle": true,
                                "wrap": true
                            }
                        ],
                        "width": "stretch"
                    }
                ]
            }
        ],
        "version": "1.5",
        "schema": "http://adaptivecards.io/schemas/adaptive-card.json"
    }
    """;

    // Deserialize the JSON into an AdaptiveCard object
    var card = AdaptiveCard.Deserialize(cardJson);

    // Send the card
    await client.Send(card);
    ```

**SDK 2.1 (current):**

```csharp
    using System.Text.Json;
    using Microsoft.Teams.Apps;
    using Microsoft.Teams.Cards;

    var cardJson = """
    {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "verticalContentAlignment": "center",
                        "items": [
                            {
                                "type": "Image",
                                "style": "Person",
                                "url": "https://aka.ms/AAp9xo4",
                                "size": "Small",
                                "altText": "Portrait of David Claux"
                            }
                        ],
                        "width": "auto"
                    },
                    {
                        "type": "Column",
                        "spacing": "medium",
                        "verticalContentAlignment": "center",
                        "items": [
                            {
                                "type": "TextBlock",
                                "weight": "Bolder",
                                "text": "David Claux",
                                "wrap": true
                            }
                        ],
                        "width": "auto"
                    },
                    {
                        "type": "Column",
                        "spacing": "medium",
                        "verticalContentAlignment": "center",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Principal Platform Architect at Microsoft",
                                "isSubtle": true,
                                "wrap": true
                            }
                        ],
                        "width": "stretch"
                    }
                ]
            }
        ],
        "version": "1.5",
        "schema": "http://adaptivecards.io/schemas/adaptive-card.json"
    }
    """;

    // Deserialize the JSON into an AdaptiveCard object
    var card = AdaptiveCard.Deserialize(cardJson);

    // Serialize and wrap in a TeamsAttachment
    JsonElement cardElement = JsonSerializer.SerializeToElement(card);
    TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(cardElement)
        .Build();

    // Send the card
    await context.SendAsync(new MessageActivityInput().AddAttachment(attachment), cancellationToken);
    ```

This method leverages the full Adaptive Card schema and ensures that the payload adheres strictly to `AdaptiveCard`.

:::tip
You can use a combination of raw JSON and builder helpers depending on whatever you find easier.
:::

## End‑to‑end Example – Task Form Card

Below is a complete example showing a task management form.

**SDK 2.0 (Legacy):**

```csharp
    using Microsoft.Teams.Apps.Annotations;

    [TeamsController]
    public class MessageController
    {
        [Message]
        public async Task OnMessage([Context] IContext context)
        {
            var text = context.Activity.Text?.ToLowerInvariant() ?? "";

            if (text.Contains("form"))
            {
                var card = CreateTaskFormCard();
                await context.Client.Send(card);
            }
        }
    }
    ```

**SDK 2.1 (current):**

```csharp
    using System.Text.Json;
    using Microsoft.Teams.Apps;
    using Microsoft.Teams.Cards;

    bot.OnMessage(async (context, cancellationToken) =>
    {
        var text = context.Activity.Text?.ToLowerInvariant() ?? "";

        if (text.Contains("form"))
        {
            await context.TypingAsync(cancellationToken);
            var card = CreateTaskFormCard();
            JsonElement cardElement = JsonSerializer.SerializeToElement(card);
            TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
                .WithAdaptiveCard(cardElement)
                .Build();
            await context.SendAsync(new MessageActivityInput().AddAttachment(attachment), cancellationToken);
        }
    });
    ```

The definition for `CreateTaskFormCard` is as follows

```csharp
private static AdaptiveCard CreateTaskFormCard()
{
    return new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock("Create New Task")
            {
                Weight = TextWeight.Bolder,
                Size = TextSize.Large
            },
            new TextInput
            {
                Id = "title",
                Label = "Task Title",
                Placeholder = "Enter task title"
            },
            new TextInput
            {
                Id = "description",
                Label = "Description",
                Placeholder = "Enter task details",
                IsMultiline = true
            },
            new ChoiceSetInput
            {
                Id = "priority",
                Label = "Priority",
                Value = "medium",
                Choices = new List
                {
                    new() { Title = "High", Value = "high" },
                    new() { Title = "Medium", Value = "medium" },
                    new() { Title = "Low", Value = "low" }
                }
            },
            new DateInput
            {
                Id = "due_date",
                Label = "Due Date",
                Value = DateTime.Now.ToString("yyyy-MM-dd")
            }
        },
        Actions = new List
        {
            new ExecuteAction
            {
                Title = "Create Task",
                Data = new Union<string, SubmitActionData>(new SubmitActionData
                {
                    NonSchemaProperties = new Dictionary<string, object?>
                    {
                        { "action", "create_task" }
                    }
                }),
                AssociatedInputs = AssociatedInputs.Auto,
                Style = ActionStyle.Positive
            }
        }
    };
}
```

## Additional Resources

- [**Official Adaptive Card Documentation**](https://adaptivecards.microsoft.com/)
- [**Adaptive Cards Designer**](https://adaptivecards.microsoft.com/designer.html)

### Summary

- Use **builder helpers** for readable, maintainable card code.
- Enjoy **full type safety** and IDE assistance.
- Prototype quickly in the **visual designer** and refine with builders.

Happy card building! 🎉
