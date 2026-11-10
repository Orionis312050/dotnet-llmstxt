---
title: "Proactive Messaging"
sidebar_position: "1"
---

# Proactive Messaging

In [Sending Messages](#), you were shown how to respond to an event when it happens. However, there are times when you want to send a message to the user without them sending a message first. This is called proactive messaging. You can do this by using the `SendAsync` method on your app instance. This approach is useful for sending notifications or reminders to the user.

The main thing to note is that you need to have the `conversationId` of the chat or channel that you want to send the message to. It's a good idea to store this value somewhere from an activity handler so that you can use it for proactive messaging later.

**SDK 2.0 (Legacy):**

```csharp
    app.OnInstall(async (context, cancellationToken) =>
    {
        // Save the conversation id in
        context.Storage.Set(context.Activity.From.AadObjectId!, context.Activity.Conversation.Id);
        await context.Send("Hi! I am going to remind you to say something to me soon!", cancellationToken);
        notificationQueue.AddReminder(context.Activity.From.AadObjectId!, Notifications.SendProactive, 10_000);
    });
    ```

**SDK 2.1+:**

```csharp
    teams.OnInstall(async (context, cancellationToken) =>
    {
        // Save the conversation id in
        context.State.UserState?.Set("conversationId", context.Activity.Conversation.Id);
        await context.SendAsync("Hi! I am going to remind you to say something to me soon!", cancellationToken);
    });
    ```

Then, when you want to send a proactive message, you can retrieve the `conversationId` from storage and use it to send the message.

**SDK 2.0 (Legacy):**

```csharp
public static class Notifications
{
    public static async Task SendProactive(string userId)
    {
        var conversationId = (string?)storage.Get(userId);

        if (conversationId is null) return;

        await app.Send(conversationId, "Hey! It's been a while. How are you?");
    }
}
```

**SDK 2.1+:**

```csharp
public static class Notifications
{
    public static async Task SendProactive(Context context, CancellationToken cancellationToken = default)
    {
        var conversationId = context.State.UserState?.Get<string>("conversationId");
        if (string.IsNullOrWhiteSpace(conversationId)) return;

        await app.SendAsync(conversationId, "Hey! It's been a while. How are you?", cancellationToken: cancellationToken);
    }
}
```

:::tip
In this example, you see how to get the `conversationId` using one of the activity handlers. This is a good place to store the conversation id, but you can also do this in other places like when the user installs the app or when they sign in. The important thing is that you have the conversation id stored somewhere so you can use it later.
:::

## Targeted Proactive Messages

Targeted messages, also known as ephemeral messages, are delivered to a specific user in a shared conversation. From a single user's perspective, they appear as regular inline messages in a conversation. Other participants won't see these messages.

When sending targeted messages proactively, you must explicitly specify the recipient account.

**SDK 2.0 (Legacy):**

```csharp
// When sending proactively, you must provide an explicit recipient account
public static async Task SendTargetedNotification(string conversationId, Account recipient)
{
    var teams = app.UseTeams();
    await teams.Send(
        conversationId,
        new MessageActivity("This is a private notification just for you!")
            .WithRecipient(recipient, isTargeted: true)
    );
}
```

**SDK 2.1+:**

```csharp
// When sending proactively, you must provide an explicit recipient account
public static async Task SendTargetedNotification(
    string conversationId,
    Account recipient,
    CancellationToken cancellationToken = default)
{
    await app.SendAsync(
        conversationId,
        new MessageActivityInput()
            .WithText("This is a private notification just for you!")
            .WithRecipient(recipient, isTargeted: true),
        cancellationToken: cancellationToken);
}
```

## Proactive Threading

The SDK sends messages into reply chains (threads) in channels and 1:1 chats.
Whether a thread is shown depends on the service and the Teams client.
The SDK sends a response from a handler through the reply chain route only when `channelData.replyChain.id` is present on the inbound activity.
With no `replyChain`, the SDK skips that route and sends to the inbound conversation.
In a channel, the suffix of that conversation id can still keep the response in the thread.

In this guide, top-level means a message in the main chat or channel conversation, and in-thread means a message in a reply chain (thread).

The `app` methods take explicit ids, so they work outside a handler.
Given the id of the chat or channel (`channelData.chat.id` or `channelData.channel.id`), in 1:1 chats and channels they place the message as follows:

| Method | 1:1 chat | Channel |
|---|---|---|
| `app.SendAsync()` | Top-level | Top-level |
| `app.ReplyAsync()` | In-thread | In-thread |

To proactively send a message as a reply in a thread, use `app.Reply()` (SDK 2.0) / `app.ReplyAsync()` (SDK 2.1+, reply chain routing in SDK 2.2) with the id of the chat or channel and the id of the top-level message of the thread.
Take the ids from `channelData` on the inbound activity:

- **Chat id:** `channelData.chat.id`, in a 1:1 chat.
- **Channel id:** `channelData.channel.id`, in a channel.
- **Reply chain id:** `channelData.replyChain.id`, the id of the top-level message of the thread.
  - A `replyChain.id` that matches the message's own `activity.Id` indicates a top-level post.
    A different `replyChain.id` indicates the message is in-thread.
  - If `replyChain` is absent, there is no reply chain id to take from the inbound activity, so pass an id you stored yourself.
    The SDK does not take it from the conversation id.

**SDK 2.0 (Legacy):**

```csharp
// SDK 2.0 builds the thread address from a channel id and the numeric id of the root message
await app.Reply(channelId, messageId, "Thread update!");
```

**SDK 2.2:**

```csharp
// chatOrChannelId is channelData.chat.id or channelData.channel.id
// replyChainId is channelData.replyChain.id
await app.ReplyAsync(chatOrChannelId, replyChainId, "Thread update!", cancellationToken: cancellationToken);
```

To send a new top-level (non-threaded) message instead, send to the id of the chat or channel with `app.Send()` (SDK 2.0) / `app.SendAsync()` (SDK >= 2.1).

**SDK 2.0 (Legacy):**

```csharp
// Send a new top-level message to a channel
await app.Send(channelId, "Hello!");
```

**SDK 2.2:**

```csharp
// Send a new top-level message to the chat or channel
await app.SendAsync(chatOrChannelId, "Hello!", cancellationToken: cancellationToken);
```

If the reply chain call fails, `app.ReplyAsync()` throws the error to your code.
It does not fall back to another route.

Inside a handler, a response goes to the `ServiceUrl` of the inbound activity.
`app.SendAsync()` and `app.ReplyAsync()` go to the `serviceUrl` argument when you pass one, and otherwise to the service URL of the last inbound activity the app received.
There is no default, so the call throws `InvalidOperationException` until the app has received an activity.

A service URL without its cloud segment can be accepted and still post the message flat, outside the thread.
Store the full service URL of the inbound activity, and pass or configure it for proactive calls, instead of rebuilding or trimming it.

The API client's reply operation, `Api.Conversations.ReplyToActivityAsync()`, is a plain send to the conversation it is given.
It posts to the activities route of that conversation and does not place the message in a reply chain.
To call the reply chain route directly, use `Api.Conversations.CreateReplyChainActivityAsync()`, which is experimental.
It takes the chat or channel id and the reply chain id, and has no fallback.

If you stored a conversation id from a channel message, it ends in `;messageid=<root post id>`.
That form identifies the thread, not the chat or channel, and the reply chain endpoint rejects it.
Store `channelData.chat.id` or `channelData.channel.id` and `channelData.replyChain.id` instead.

A proactive call has no inbound activity to take an identity from.
When the recipient is an agentic user, pass its identity from the inbound activity's recipient.
Without it, the call uses the app's own credentials instead of the agentic user's.
An agentic app cannot get an app-only token, so the call fails with an authentication error (`AADSTS82001`) and nothing is sent.

```csharp
// SDK 2.2 and later. AgenticIdentity is null unless the recipient is an agentic user
await app.ReplyAsync(
    chatOrChannelId,
    replyChainId,
    "Thread update!",
    agenticIdentity: context.Activity.Recipient?.GetAgenticIdentity(),
    cancellationToken: cancellationToken);
```

For reactive threading (within a handler), see [Threading](#threading).
