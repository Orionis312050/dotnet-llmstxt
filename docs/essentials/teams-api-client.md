---
title: "Teams API Client"
sidebar_position: "6"
---

# Teams API Client

Teams has a number of areas that your application has access to via its API. These are all available via the `context.Api` (and `teams.Api` at the app level) object. Here is a short summary of the different areas:

| Area            | Description                                                                                                                                                          |
| --------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Conversations` | Gives your application the ability to perform activities on conversations (send, update, delete messages, etc.), or create conversations (like 1:1 chat with a user) |
| `Meetings`      | Gives your application access to meeting details and participant information via `GetByIdAsync` and `GetParticipantAsync`                                             |
| `Teams`         | Gives your application access to team or channel details                                                                                                             |

An instance of the API client is passed to handlers that can be used to fetch details:

## Example

In this example, we use the API client to fetch the members in a conversation. The `Api` object is passed to the activity handler in this case.

**SDK 2.0 (Legacy):**

```csharp
app.OnMessage(async (context, cancellationToken) =>
{
    var members = await context.Api.Conversations.Members.Get(context.Conversation.Id);
});
```

**SDK 2.1 (current):**

```csharp
teams.OnMessage(async (context, cancellationToken) =>
{
    var members = await context.Api.Conversations.Members.GetAsync(
        context.Activity.Conversation!.Id!,
        cancellationToken: cancellationToken);
});
```

## Proactive API

It's also possible to access the API client from outside a handler via the app instance. Here we have the same example as above, but we're access the API client via the app instance.

**SDK 2.0 (Legacy):**

```csharp
var members = await app.Api.Conversations.Members.Get("...");
```

**SDK 2.1 (current):**

Outside of a turn, use the app-level `teams.Api` client (point it at the target service URL first):

```csharp
var api = teams.Api.ForServiceUrl(serviceUrl);
var members = await api.Conversations.Members.GetAsync(conversationId);
```

## Meetings Example

In this example, we use the API client to get a specific meeting participant's details, such as their role (e.g. Organizer) and whether they are currently in the meeting. Provide the user's AAD Object ID to specify which participant to look up. The `meetingId` and `tenantId` are available from the activity's channel data.

:::note
To retrieve **all** members of a meeting, use the conversations API as shown in the [example above](#example), since meetings are also conversations.
:::

**SDK 2.0 (Legacy):**

```csharp
app.OnMeetingStart(async (context, cancellationToken) =>
{
    var meetingId = context.Activity.Value.Id;
    var tenantId = context.Activity.ChannelData?.Tenant?.Id;
    var userId = context.Activity.From?.AadObjectId;

    if (meetingId != null && tenantId != null && userId != null)
    {
        var participant = await context.Api.Meetings.GetParticipantAsync(meetingId, userId, tenantId);
        // participant.Meeting?.Role — "Organizer", "Presenter", "Attendee"
        // participant.Meeting?.InMeeting — true/false
    }
});
```

**SDK 2.1 (current):**

```csharp
teams.OnMeetingStart(async (context, cancellationToken) =>
{
    var meetingId = context.Activity.Value.Id;
    var tenantId = context.Activity.ChannelData?.Tenant?.Id;
    var userId = context.Activity.From?.AadObjectId;

    if (meetingId != null && tenantId != null && userId != null)
    {
        var participant = await context.Api.Meetings.GetParticipantAsync(meetingId, userId, tenantId);
        // participant.Meeting?.Role — "Organizer", "Presenter", "Attendee"
        // participant.Meeting?.InMeeting — true/false
    }
});
```

Visit [Meeting Events](../in-depth-guides/meeting-events.md) to learn more about meeting events.
