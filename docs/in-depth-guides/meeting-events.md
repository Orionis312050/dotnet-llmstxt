---
title: "Meeting Events"
sidebar_position: "5"
---

# Meeting Events

Microsoft Teams provides meeting events that allow your application to respond to various meeting lifecycle changes. Your app can listen to events like when a meeting starts, meeting ends, and participant activities to create rich, interactive experiences.

## Overview

Meeting events enable your application to:
- Send notifications when meetings start or end
- Track participant activity (join/leave events)
- Display relevant information or cards based on meeting context
- Integrate with meeting workflows

## Configuring Your Bot

There are a few requirements in the Teams app manifest (`manifest.json`) to support these events.

1. The scopes section must include `team`, and `groupChat`

```json
bots": [
        {
            "botId": "",
            "scopes": [
                "team",
                "personal",
                "groupChat"
            ],
            "isNotificationOnly": false
        }
    ]
```

2. In the authorization section, make sure to specify the following resource-specific permissions:

```json
 "authorization":{
        "permissions":{
            "resourceSpecific":[
                {
                    "name":"OnlineMeetingParticipant.Read.Chat",
                    "type":"Application"
                },
                {
                    "name":"ChannelMeeting.ReadBasic.Group",
                    "type":"Application"
                },
                {
                    "name":"OnlineMeeting.ReadBasic.Chat",
                    "type":"Application"
                }
                ]
            }
        }
```

3. In the Teams Developer Portal, for your `Bot`, make sure the `Meeting Event Subscriptions` are checked off. This enables you to receive the Meeting Participant events. For these events, you must create your Bot via TDP.

## Meeting Start Event

When a meeting starts, your app can handle the `meetingStart` event to send a notification or card to the meeting chat.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register meeting start handler
app.OnMeetingStart(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var startTime = activity.StartTime.ToLocalTime();

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"'{activity.Title}' has started at {startTime}.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        },
        // Meetings held inside a channel have no join URL, so there is no action to offer.
        Actions = activity.JoinUrl is null
            ? new List()
            : new List
            {
                new OpenUrlAction(activity.JoinUrl)
                {
                    Title = "Join the meeting",
                }
            }
    };

    await context.Send(card, cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register meeting start handler
teams.OnMeetingStart(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var startTime = DateTimeOffset.Parse(activity.StartTime).ToLocalTime();

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"'{activity.Title}' has started at {startTime}.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        },
        // Meetings held inside a channel have no join URL, so there is no action to offer.
        Actions = activity.JoinUrl is null
            ? new List()
            : new List
            {
                new OpenUrlAction(activity.JoinUrl)
                {
                    Title = "Join the meeting",
                }
            }
    };

    TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(JsonSerializer.SerializeToElement(card))
        .Build();
    await context.SendAsync(new MessageActivityInput().AddAttachment(attachment), cancellationToken);
});
```

## Meeting End Event

When a meeting ends, your app can handle the `meetingEnd` event to send a summary or follow-up information.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register meeting end handler
app.OnMeetingEnd(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var endTime = activity.EndTime.ToLocalTime();

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"'{activity.Title}' has ended at {endTime}.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        }
    };

    await context.Send(card, cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register meeting end handler
teams.OnMeetingEnd(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var endTime = DateTimeOffset.Parse(activity.EndTime).ToLocalTime();

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"'{activity.Title}' has ended at {endTime}.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        }
    };

    TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(JsonSerializer.SerializeToElement(card))
        .Build();
    await context.SendAsync(new MessageActivityInput().AddAttachment(attachment), cancellationToken);
});
```

## Participant Join Event

When a participant joins a meeting, your app can handle the `meetingParticipantJoin` event to welcome them or display their role.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register participant join handler
app.OnMeetingJoin(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var member = activity.Members[0].User.Name;
    var role = activity.Members[0].Meeting.Role;

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"{member} has joined the meeting as {role}.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        }
    };

    await context.Send(card, cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register participant join handler
teams.OnMeetingJoin(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var member = activity.Members[0].User.Name;
    var role = activity.Members[0].Meeting.Role;

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"{member} has joined the meeting as {role}.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        }
    };

    TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(JsonSerializer.SerializeToElement(card))
        .Build();
    await context.SendAsync(new MessageActivityInput().AddAttachment(attachment), cancellationToken);
});
```

## Participant Leave Event

When a participant leaves a meeting, your app can handle the `meetingParticipantLeave` event to notify others.

**SDK 2.0 (Legacy):**

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register participant leave handler
app.OnMeetingLeave(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var member = activity.Members[0].User.Name;

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"{member} has left the meeting.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        }
    };

    await context.Send(card, cancellationToken);
});
```

**SDK 2.1 (current):**

```csharp
using System.Text.Json;
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Activities;
using Microsoft.Teams.Apps.Activities.Events;
using Microsoft.Teams.Cards;

// Register participant leave handler
teams.OnMeetingLeave(async (context, cancellationToken) =>
{
    var activity = context.Activity.Value;
    var member = activity.Members[0].User.Name;

    var card = new AdaptiveCard
    {
        Schema = "http://adaptivecards.io/schemas/adaptive-card.json",
        Body = new List
        {
            new TextBlock($"{member} has left the meeting.")
            {
                Wrap = true,
                Weight = TextWeight.Bolder
            }
        }
    };

    TeamsAttachment attachment = TeamsAttachment.CreateBuilder()
        .WithAdaptiveCard(JsonSerializer.SerializeToElement(card))
        .Build();
    await context.SendAsync(new MessageActivityInput().AddAttachment(attachment), cancellationToken);
});
```
