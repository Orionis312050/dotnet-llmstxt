---
title: "Creating Dialogs"
sidebar_position: "1"
---

# Creating Dialogs

:::tip
If you're not familiar with how to build Adaptive Cards, check out [the cards guide](../adaptive-cards.md). Understanding their basics is a prerequisite for this guide.
:::

## Entry Point

To open a dialog, you need to supply a special type of action to the Adaptive Card. The `TaskFetchAction` is specifically designed for this purpose - it automatically sets up the proper Teams data structure to trigger a dialog. Once this button is clicked, the dialog will open and ask the application what to show.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Api.Activities;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Annotations;
using Microsoft.Teams.Cards;
using Microsoft.Teams.Common.Logging;

//...

[Message]
public async Task OnMessage([Context] MessageActivity activity, [Context] IContext.Client client, [Context] ILogger log)
{
    // Create the launcher adaptive card
    var card = CreateDialogLauncherCard();
    await client.Send(card);
}

private static AdaptiveCard CreateDialogLauncherCard()
{
    var card = new AdaptiveCard
    {
        Body = new List
        {
            new TextBlock("Select the examples you want to see!")
            {
                Size = TextSize.Large,
                Weight = TextWeight.Bolder
            }
        },
        Actions = new List
        {
            new TaskFetchAction(new { opendialogtype = "simple_form" })
            {
                Title = "Simple form test"
            },
            new TaskFetchAction(new { opendialogtype = "webpage_dialog" })
            {
                Title = "Webpage Dialog"
            },
            new TaskFetchAction(new { opendialogtype = "multi_step_form" })
            {
                Title = "Multi-step Form"
            }
        }
    };

    return card;
}
```

**SDK 2.1 (current):**

```csharp
using Microsoft.Teams.Cards;

//...

teams.OnMessage(async (context, cancellationToken) =>
{
    // Create the launcher adaptive card
    var card = CreateDialogLauncherCard();
    await context.SendAsync(card, cancellationToken);
});

private static AdaptiveCard CreateDialogLauncherCard()
{
    var card = new AdaptiveCard
    {
        Body = new List
        {
            new TextBlock("Select the examples you want to see!")
            {
                Size = TextSize.Large,
                Weight = TextWeight.Bolder
            }
        },
        Actions = new List
        {
            new TaskFetchAction(new { opendialogtype = "simple_form" })
            {
                Title = "Simple form test"
            },
            new TaskFetchAction(new { opendialogtype = "webpage_dialog" })
            {
                Title = "Webpage Dialog"
            },
            new TaskFetchAction(new { opendialogtype = "multi_step_form" })
            {
                Title = "Multi-step Form"
            }
        }
    };

    return card;
}
```

## Handling Dialog Open Events

Once an action is executed to open a dialog, the Teams client will send an event to the agent to request what the content of the dialog should be. When using `TaskFetchAction`, the data is nested inside an `MsTeams` property structure.

**SDK 2.0 (Legacy):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Api.TaskModules;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities.Invokes;
using Microsoft.Teams.Apps.Annotations;
using Microsoft.Teams.Common.Logging;

//...

[TaskFetch]
public Microsoft.Teams.Api.TaskModules.Response OnTaskFetch([Context] Tasks.FetchActivity activity, [Context] IContext.Client client, [Context] ILogger log)
{
    var data = activity.Value?.Data as JsonElement?;
    if (data == null)
    {
        log.Info("[TASK_FETCH] No data found in the activity value");
        return new Microsoft.Teams.Api.TaskModules.Response(
            new Microsoft.Teams.Api.TaskModules.MessageTask("No data found in the activity value"));
    }

    var dialogType = data.Value.TryGetProperty("opendialogtype", out var dialogTypeElement) && dialogTypeElement.ValueKind == JsonValueKind.String
        ? dialogTypeElement.GetString()
        : null;

    log.Info($"[TASK_FETCH] Dialog type: {dialogType}");

    return dialogType switch
    {
        "simple_form" => CreateSimpleFormDialog(),
        "webpage_dialog" => CreateWebpageDialog(_configuration, log),
        "multi_step_form" => CreateMultiStepFormDialog(),
        "mixed_example" => CreateMixedExampleDialog(),
        _ => new Microsoft.Teams.Api.TaskModules.Response(
            new Microsoft.Teams.Api.TaskModules.MessageTask("Unknown dialog type"))
    };
}
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps.TaskModules;

//...

teams.OnTaskFetch(async (context, cancellationToken) =>
{
    var data = context.Activity.Value?.Data as JsonElement?;
    if (data == null)
    {
        return TaskModuleResponse.CreateBuilder()
            .WithType(TaskModuleResponseTypes.Message)
            .WithMessage("No data found in the activity value")
            .Build();
    }

    var dialogType = data.Value.TryGetProperty("opendialogtype", out var dialogTypeElement) && dialogTypeElement.ValueKind == JsonValueKind.String
        ? dialogTypeElement.GetString()
        : null;

    return dialogType switch
    {
        "simple_form" => CreateSimpleFormDialog(),
        "webpage_dialog" => CreateWebpageDialog(_configuration, context.Log),
        "multi_step_form" => CreateMultiStepFormDialog(),
        "mixed_example" => CreateMixedExampleDialog(),
        _ => TaskModuleResponse.CreateBuilder()
            .WithType(TaskModuleResponseTypes.Message)
            .WithMessage("Unknown dialog type")
            .Build()
    };
});
```

### Rendering A Card

You can render an Adaptive Card in a dialog by returning a card response.

**SDK 2.0 (Legacy):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Api;
using Microsoft.Teams.Api.TaskModules;
using Microsoft.Teams.Cards;

//...

private static Microsoft.Teams.Api.TaskModules.Response CreateSimpleFormDialog()
{
    var choices = new List
    {
        new Choice { Title = "Option 1", Value = "opt1" },
        new Choice { Title = "Option 2", Value = "opt2" },
        new Choice { Title = "Option 3", Value = "opt3" }
    };

    var dialogCard = new AdaptiveCard
    {
        Body = new List
        {
            new TextBlock("This is a simple form")
            {
                Size = TextSize.Large,
                Weight = TextWeight.Bolder
            },
            new TextInput
            {
                Id = "name",
                Label = "Name",
                Placeholder = "Enter your name",
                IsRequired = true
            },
            new ChoiceSetInput
            {
                Id = "preference",
                Label = "Select your preference",
                Choices = choices,
                Style = StyleEnum.Compact
            }
        },
        Actions = new List
        {
            new SubmitAction
            {
                Title = "Submit",
                Data = new { submissiondialogtype = "simple_form" }
            }
        }
    };

    var taskInfo = new TaskInfo
    {
        Title = "Simple Form Dialog",
        Card = new Attachment
        {
            ContentType = new ContentType("application/vnd.microsoft.card.adaptive"),
            Content = dialogCard
        }
    };

    return new Microsoft.Teams.Api.TaskModules.Response(
        new Microsoft.Teams.Api.TaskModules.ContinueTask(taskInfo));
}
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.TaskModules;
using Microsoft.Teams.Cards;

//...

private static TaskModuleResponse CreateSimpleFormDialog()
{
    var choices = new List
    {
        new Choice { Title = "Option 1", Value = "opt1" },
        new Choice { Title = "Option 2", Value = "opt2" },
        new Choice { Title = "Option 3", Value = "opt3" }
    };

    var dialogCard = new AdaptiveCard
    {
        Body = new List
        {
            new TextBlock("This is a simple form")
            {
                Size = TextSize.Large,
                Weight = TextWeight.Bolder
            },
            new TextInput
            {
                Id = "name",
                Label = "Name",
                Placeholder = "Enter your name",
                IsRequired = true
            },
            new ChoiceSetInput
            {
                Id = "preference",
                Label = "Select your preference",
                Choices = choices,
                Style = StyleEnum.Compact
            }
        },
        Actions = new List
        {
            new SubmitAction
            {
                Title = "Submit",
                Data = new { submissiondialogtype = "simple_form" }
            }
        }
    };

    TeamsAttachment taskModuleCardResponse = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(JsonSerializer.SerializeToElement(dialogCard))
        .Build();

    return TaskModuleResponse.CreateBuilder()
        .WithType(TaskModuleResponseTypes.Continue)
        .WithTitle("Simple Form Dialog")
        .WithHeight("medium")
        .WithWidth("medium")
        .WithCard(taskModuleCardResponse)
        .Build();
}
```

:::info
The action type for submitting a dialog must be `Action.Submit`. This is a requirement of the Teams client. If you use a different action type, the dialog will not be submitted and the agent will not receive the submission event.
:::

### Rendering A Webpage

You can render a webpage in a dialog as well. There are some security requirements to be aware of:

1. The webpage must be hosted on a domain that is allow-listed as `validDomains` in the Teams app [manifest](../../teams/teams-manifest.md) for the agent
2. The webpage must also host the [teams-js client library](https://www.npmjs.com/package/@microsoft/teams-js). The reason for this is that for security purposes, the Teams client will not render arbitrary webpages. As such, the webpage must explicitly opt-in to being rendered in the Teams client. Setting up the teams-js client library handles this for you.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Api.TaskModules;
using Microsoft.Teams.Common;

//...

private static Microsoft.Teams.Api.TaskModules.Response CreateWebpageDialog(IConfiguration configuration, ILogger log)
{
    var botEndpoint = configuration["BotEndpoint"];
    if (string.IsNullOrEmpty(botEndpoint))
    {
        botEndpoint = "http://localhost:3978"; // Fallback for local development
    }

    var taskInfo = new TaskInfo
    {
        Title = "Webpage Dialog",
        Width = new Union<int, Size>(1000),
        Height = new Union<int, Size>(800),
        // Here we are using a webpage that is hosted in the same
        // server as the agent. This server needs to be publicly accessible,
        // needs to set up teams.js client library (https://www.npmjs.com/package/@microsoft/teams-js)
        // and needs to be registered in the manifest.
        Url = $"{botEndpoint}/tabs/dialog-form"
    };

    return new Microsoft.Teams.Api.TaskModules.Response(
        new Microsoft.Teams.Api.TaskModules.ContinueTask(taskInfo));
}
```

**SDK 2.1 (current):**

```csharp
using Microsoft.Teams.Apps.Schema;
using Microsoft.Teams.Apps.TaskModules;

//...

private static InvokeResponse CreateWebpageDialog(IConfiguration configuration)
{
    var botEndpoint = configuration["BotEndpoint"] ?? "http://localhost:3978";

    // The TaskModuleResponse builder supports card-based dialogs only.
    // For URL-based dialogs, construct the response manually.
    // This server needs to be publicly accessible, set up the teams.js client library
    // (https://www.npmjs.com/package/@microsoft/teams-js), and be registered in the manifest.
    return new InvokeResponse(200, new TaskModuleResponse
    {
        Task = new Microsoft.Teams.Apps.TaskModules.Response
        {
            Type = TaskModuleResponseTypes.Continue,
            Value = new
            {
                title = "Webpage Dialog",
                url = $"{botEndpoint}/tabs/dialog-form",
                height = 800,
                width = 1000
            }
        }
    });
}
```

### Setting up Embedded Web Content

To serve web content for dialogs, you can use the `AddTab` functionality to embed HTML files as resources:

```csharp
// In Program.cs when building your app
app.UseTeams();
app.AddTab("dialog-form", "Web/dialog-form");

// Configure project file to embed web resources
// In .csproj:
// true
// 
// 
```
