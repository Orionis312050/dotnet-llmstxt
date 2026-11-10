---
title: "Executing Actions"
sidebar_position: "2"
---

# Executing Actions

Adaptive Cards support interactive elements through **actions**—buttons, links, and input submission triggers that respond to user interaction.
You can use these to collect form input, trigger workflows, show task modules, open URLs, and more.

## Action Types

The Teams SDK supports several action types for different interaction patterns:

| Action Type               | Purpose                | Description                                                                  |
| ------------------------- | ---------------------- | ---------------------------------------------------------------------------- |
| `Action.Execute`          | Server‑side processing | Send data to your bot for processing. Best for forms & multi‑step workflows. |
| `Action.Submit`           | Simple data submission | Legacy action type. Prefer `Execute` for new projects.                       |
| `Action.OpenUrl`          | External navigation    | Open a URL in the user's browser.                                            |
| `Action.ShowCard`         | Progressive disclosure | Display a nested card when clicked.                                          |
| `Action.ToggleVisibility` | UI state management    | Show/hide card elements dynamically.                                         |

:::info
For complete reference, see the [official documentation](https://adaptivecards.microsoft.com/?topic=Action.Execute).
:::

## Creating Actions with the SDK

### Single Actions

The SDK provides builder helpers that abstract the underlying JSON. For example:

```csharp
using Microsoft.Teams.Cards;

var action = new ExecuteAction
{
    Title = "Submit Feedback",
    Data = new Union<string, SubmitActionData>(new SubmitActionData
    {
        NonSchemaProperties = new Dictionary<string, object?>
        {
            { "action", "submit_feedback" }
        }
    }),
    AssociatedInputs = AssociatedInputs.Auto
};
```

### Action Sets

Group actions together using `ActionSet`:

```csharp
using Microsoft.Teams.Cards;

var card = new AdaptiveCard
{
    Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
    Actions = new List
    {
        new ExecuteAction
        {
            Title = "Submit Feedback",
            Data = new Union<string, SubmitActionData>(new SubmitActionData
            {
                NonSchemaProperties = new Dictionary<string, object?>
                {
                    { "action", "submit_feedback" }
                }
            })
        },
        new OpenUrlAction("https://adaptivecards.microsoft.com")
        {
            Title = "Learn More"
        }
    }
};
```

### Raw JSON Alternative

Just like when building cards, if you prefer to work with raw JSON, you can do just that. 

```csharp
var actionJson = """
{
  "type": "Action.OpenUrl",
  "url": "https://adaptivecards.microsoft.com",
  "title": "Learn More"
}
""";
var action = OpenUrlAction.Deserialize(actionJson);
```

## Working with Input Values

### Associating data with the cards

Sometimes you want to send a card and have it be associated with some data. Set the `data` value to be sent back to the client so you can associate it with a particular entity.

```csharp
private static AdaptiveCard CreateProfileCard()
{
    return new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock("User Profile")
            {
                Weight = TextWeight.Bolder,
                Size = TextSize.Large
            },
            new TextInput
            {
                Id = "name",
                Label = "Name",
                Value = "John Doe"
            },
            new TextInput
            {
                Id = "email",
                Label = "Email",
                Value = "john@contoso.com"
            },
            new ToggleInput("Subscribe to newsletter")
            {
                Id = "subscribe",
                Value = "false"
            }
        },
        Actions = new List
        {
            new ExecuteAction
            {
                Title = "Save",
                // entity_id will come back after the user submits
                Data = new Union<string, SubmitActionData>(new SubmitActionData
                {
                    NonSchemaProperties = new Dictionary<string, object?>
                    {
                        { "action", "save_profile" },
                        { "entity_id", "12345" }
                    }
                }),
                AssociatedInputs = AssociatedInputs.Auto
            }
        }
    };
}

// Data received in handler (conceptual structure)
/*
{
  "action": "save_profile",
  "entity_id": "12345",     // From action data
  "name": "John Doe",       // From name input
  "email": "john@doe.com",  // From email input
  "subscribe": "true"       // From toggle input (as string)
}

Accessed in C# as:
- data["action"] → "save_profile"
- data["entity_id"] → "12345"
- data["name"] → "John Doe"
- data["email"] → "john@doe.com"
- data["subscribe"] → "true"
*/
```

### Input Validation

Input Controls provide ways for you to validate. More details can be found on the Adaptive Cards [documentation](https://adaptivecards.microsoft.com/?topic=input-validation).

```csharp
private static AdaptiveCard CreateProfileCardWithValidation()
{
    return new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock("Profile with Validation")
            {
                Weight = TextWeight.Bolder,
                Size = TextSize.Large
            },
            new NumberInput
            {
                Id = "age",
                Label = "Age",
                IsRequired = true,
                Min = 0,
                Max = 120
            },
            // Can configure custom error messages
            new TextInput
            {
                Id = "name",
                Label = "Name",
                IsRequired = true,
                ErrorMessage = "Name is required"
            },
            new TextInput
            {
                Id = "location",
                Label = "Location"
            }
        },
        Actions = new List
        {
            new ExecuteAction
            {
                Title = "Save",
                // All inputs should be validated
                Data = new Union<string, SubmitActionData>(new SubmitActionData
                {
                    NonSchemaProperties = new Dictionary<string, object?>
                    {
                        { "action", "save_profile" }
                    }
                }),
                AssociatedInputs = AssociatedInputs.Auto
            }
        }
    };
}
```

## Server Handlers

### Basic Structure

Card actions arrive as `card.action` activities in your app. These give you access to the validated input values plus any `data` values you had configured to be sent back to you.

**SDK 2.0 (Legacy):**

```csharp
    using System.Text.Json;
    using Microsoft.Teams.Api.AdaptiveCards;
    using Microsoft.Teams.Apps.Annotations;

    [TeamsController]
    public class AdaptiveCardController
    {
        [AdaptiveCard.Action]
        public async Task OnAction([Context] IContext context, CancellationToken cancellationToken)
        {
            var data = context.Activity.Value?.Action?.Data;

            if (data == null)
                return new ActionResponse.Message("No data specified") { StatusCode = 400 };

            string? action = data.TryGetValue("action", out var actionObj) ? actionObj?.ToString() : null;

            if (string.IsNullOrEmpty(action))
                return new ActionResponse.Message("No action specified") { StatusCode = 400 };

            string? GetFormValue(string key)
            {
                if (data.TryGetValue(key, out var val))
                {
                    if (val is JsonElement element)
                        return element.GetString();
                    return val?.ToString();
                }
                return null;
            }

            switch (action)
            {
                case "submit_basic":
                    var notifyValue = GetFormValue("notify") ?? "false";
                    await context.Client.Send($"Basic card submitted! Notify setting: {notifyValue}");
                    break;

                case "submit_feedback":
                    var feedbackText = GetFormValue("feedback") ?? "No feedback provided";
                    await context.Client.Send($"Feedback received: {feedbackText}");
                    break;

                case "create_task":
                    var title = GetFormValue("title") ?? "Untitled";
                    var priority = GetFormValue("priority") ?? "medium";
                    var dueDate = GetFormValue("due_date") ?? "No date";
                    await context.Client.Send($"Task created!\nTitle: {title}\nPriority: {priority}\nDue: {dueDate}");
                    break;

                case "save_profile":
                    var name = GetFormValue("name") ?? "Unknown";
                    var email = GetFormValue("email") ?? "No email";
                    var subscribe = GetFormValue("subscribe") ?? "false";
                    var age = GetFormValue("age");
                    var location = GetFormValue("location") ?? "Not specified";

                    var response = $"Profile saved!\nName: {name}\nEmail: {email}\nSubscribed: {subscribe}";
                    if (!string.IsNullOrEmpty(age))
                        response += $"\nAge: {age}";
                    if (location != "Not specified")
                        response += $"\nLocation: {location}";

                    await context.Client.Send(response);
                    break;

                case "test_json":
                    await context.Client.Send("JSON deserialization test successful!");
                    break;

                default:
                    return new ActionResponse.Message("Unknown action") { StatusCode = 400 };
            }

            return new ActionResponse.Message("Action processed successfully");
        }
    }
    ```

**SDK 2.1 (current):**

```csharp
    using System.Text.Json;
    using Microsoft.Teams.Apps;

    bot.OnAdaptiveCardAction(async (context, cancellationToken) =>
    {
        var data = context.Activity.Value?.Action?.Data;

        if (data == null)
            return AdaptiveCardResponse.CreateMessageResponse("No data specified", 400);

        string? action = data.TryGetValue("action", out var actionObj) ? actionObj?.ToString() : null;

        if (string.IsNullOrEmpty(action))
            return AdaptiveCardResponse.CreateMessageResponse("No action specified", 400);

        string? GetFormValue(string key)
        {
            if (data.TryGetValue(key, out var val))
            {
                if (val is JsonElement element)
                    return element.GetString();
                return val?.ToString();
            }
            return null;
        }

        switch (action)
        {
            case "submit_basic":
                var notifyValue = GetFormValue("notify") ?? "false";
                await context.SendAsync($"Basic card submitted! Notify setting: {notifyValue}", cancellationToken);
                break;

            case "submit_feedback":
                var feedbackText = GetFormValue("feedback") ?? "No feedback provided";
                await context.SendAsync($"Feedback received: {feedbackText}", cancellationToken);
                break;

            case "create_task":
                var title = GetFormValue("title") ?? "Untitled";
                var priority = GetFormValue("priority") ?? "medium";
                var dueDate = GetFormValue("due_date") ?? "No date";
                await context.SendAsync($"Task created!\nTitle: {title}\nPriority: {priority}\nDue: {dueDate}", cancellationToken);
                break;

            case "save_profile":
                var name = GetFormValue("name") ?? "Unknown";
                var email = GetFormValue("email") ?? "No email";
                var subscribe = GetFormValue("subscribe") ?? "false";
                var age = GetFormValue("age");
                var location = GetFormValue("location") ?? "Not specified";

                var response = $"Profile saved!\nName: {name}\nEmail: {email}\nSubscribed: {subscribe}";
                if (!string.IsNullOrEmpty(age))
                    response += $"\nAge: {age}";
                if (location != "Not specified")
                    response += $"\nLocation: {location}";

                await context.SendAsync(response, cancellationToken);
                break;

            case "test_json":
                await context.SendAsync("JSON deserialization test successful!", cancellationToken);
                break;

            default:
                return AdaptiveCardResponse.CreateMessageResponse("Unknown action", 400);
        }

        return AdaptiveCardResponse.CreateMessageResponse("Action processed successfully");
    });
    ```

:::note
The `data` values come from JSON and need to be extracted using the helper method shown above to handle different JSON element types.
:::

## Responding from a card action

A message sent from a card action handler is placed relative to the card that was clicked.
In a 1:1 chat, when the card is inside a thread, `context.SendAsync()` sends the message into that thread, and when the card is at the top level it sends the message at the top level.
This applies to `Action.Execute` and `Action.Submit`.
In a channel, the response to a card button that arrives as a `message` (`Action.Submit`) is sent through the thread's reply chain, and the SDK sends the response to any other card action to the inbound conversation.
See [Responding to edits, reactions and card actions](../../essentials/sending-messages.md#responding-to-edits-reactions-and-card-actions) for the rules.

## Dynamic Typeahead Search

An `Input.ChoiceSet` can fetch its choices dynamically as the user types, instead of listing them statically. Leave `choices` as an empty array and add a `choices.data` `Data.Query` that names a `dataset`. As the user types, Teams sends your bot an `application/search` invoke, and you return the matching results.

```csharp
using Microsoft.Teams.Cards;

private static AdaptiveCard CreateDynamicSearchCard()
{
    return new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new ChoiceSetInput
            {
                Id = "game",
                Label = "Game",
                Placeholder = "Search for a game",
                Style = ChoiceSetInputStyle.Filtered,
                Choices = new List()
            }.WithChoicesData(new QueryData { Dataset = "nintendoGames" })
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
                        { "action", "submit_game" }
                    }
                })
            }
        }
    };
}
```

Register a handler for the `application/search` invoke to return results as ` title, value ` pairs, where `title` is the display text and `value` is what gets submitted when the choice is selected:

**SDK 2.0 (Legacy):**

:::note
    Dynamic typeahead search for Adaptive Cards is not supported in SDK 2.0. Use SDK 2.1 for this feature.
    :::

**SDK 2.1 (current):**

```csharp
    using Microsoft.Teams.Apps;
    using Microsoft.Teams.Apps.Handlers;

    var games = new[] { "Super Mario Odyssey", "Metroid Dread", "Splatoon 3" };

    bot.OnSearch((context, cancellationToken) =>
    {
        var query = context.Activity.Value?.QueryText?.ToLowerInvariant() ?? "";
        var results = games
            .Where(g => g.ToLowerInvariant().Contains(query))
            .Select(g => new SearchResult { Title = g, Value = g })
            .ToList();

        var response = new SearchResponse
        {
            Value = new SearchResponseValue { Results = results }
        };

        return Task.FromResult(new InvokeResponse(200, response));
    });
    ```

The activity value also carries the query options (`skip`/`top` for paging) and the originating `dataset`, so a single handler can branch on the dataset when several `Input.ChoiceSet` inputs share it.
