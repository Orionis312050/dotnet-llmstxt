---
title: "Message Reactions"
sidebar_position: "6"
---

# Message Reactions

![Reaction emoji sent by bot](/screenshots/reaction-emoji.png)

Message reactions allow your agent to add or remove emoji reactions on messages in a Teams conversation, and to react to reactions added by users. This gives your agent a quick, low-friction way to acknowledge messages or signal status without sending a full text reply.

## Adding a Reaction

To add a reaction to a message, use the reactions client through the API client:

**SDK 2.0 (Legacy):**

```csharp
    using Microsoft.Teams.Apps.Annotations;

    [TeamsController]
    public class MessageController
    {
        [Message]
        public async Task OnMessage([Context] IContext context, CancellationToken cancellationToken)
        {
            await context.Client.Send("Hello! I'll react to your message.");

            await context.Api.Conversations.Reactions.AddAsync(
                context.Activity.Conversation.Id,
                context.Activity.Id,
                ReactionType.Like,
                cancellationToken
            );
        }
    }
    ```

**SDK 2.1 (current):**

```csharp
    using Microsoft.Teams.Apps;

    app.OnMessage(async (context, cancellationToken) =>
    {
        await context.SendAsync("Hello! I'll react to your message.", cancellationToken);

        // Add a reaction to the incoming message
        await context.Api.Conversations.AddReactionAsync(
            context.Activity.Conversation.Id,
            context.Activity.Id,
            ReactionType.Like,
            cancellationToken: cancellationToken
        );
    });
    ```

## Removing a Reaction

You can also remove reactions that your agent has previously added:

**SDK 2.0 (Legacy):**

```csharp
    using Microsoft.Teams.Apps.Annotations;

    [TeamsController]
    public class MessageController
    {
        [Message]
        public async Task OnMessage([Context] IContext context, CancellationToken cancellationToken)
        {
            // First, add a reaction
            await context.Api.Conversations.Reactions.AddAsync(
                context.Activity.Conversation.Id,
                context.Activity.Id,
                ReactionType.Heart,
                cancellationToken
            );

            // Wait a bit, then remove it
            await Task.Delay(2000, cancellationToken);
            await context.Api.Conversations.Reactions.DeleteAsync(
                context.Activity.Conversation.Id,
                context.Activity.Id,
                ReactionType.Heart,
                cancellationToken
            );
        }
    }
    ```

**SDK 2.1 (current):**

```csharp
    using Microsoft.Teams.Apps;

    app.OnMessage(async (context, cancellationToken) =>
    {
        // First, add a reaction
        await context.Api.Conversations.AddReactionAsync(
            context.Activity.Conversation.Id,
            context.Activity.Id,
            ReactionType.Heart,
            cancellationToken: cancellationToken
        );

        // Wait a bit, then remove it
        await Task.Delay(2000, cancellationToken);
        await context.Api.Conversations.DeleteReactionAsync(
            context.Activity.Conversation.Id,
            context.Activity.Id,
            ReactionType.Heart,
            cancellationToken: cancellationToken
        );
    });
    ```

## Receiving Reactions

Your agent can also listen for reactions that users add or remove on messages your agent sent.
In 1:1 chats and channels, reactions on messages your agent did not send are not delivered.
The activity carries `ReactionsAdded` and `ReactionsRemoved` collections.
Its `ReplyToId` is the id of the message that was reacted to.

.NET exposes a single `OnMessageReaction` handler plus dedicated `OnMessageReactionAdded` / `OnMessageReactionRemoved` sub-handlers.

**SDK 2.0 (Legacy):**

```csharp
    using Microsoft.Teams.Apps.Annotations;

    [TeamsController]
    public class ReactionController
    {
        [Message.Reaction]
        public void OnMessageReaction([Context] IContext context)
        {
            foreach (var reaction in context.Activity.ReactionsAdded ?? [])
            {
                Console.WriteLine($"User added reaction: {reaction.Type}");
            }

            foreach (var reaction in context.Activity.ReactionsRemoved ?? [])
            {
                Console.WriteLine($"User removed reaction: {reaction.Type}");
            }
        }
    }
    ```

**SDK 2.1 (current):**

```csharp
    using Microsoft.Teams.Apps;

    app.OnMessageReactionAdded(async (context, cancellationToken) =>
    {
        foreach (var reaction in context.Activity.ReactionsAdded ?? [])
        {
            Console.WriteLine($"User added reaction: {reaction.Type}");
        }
    });

    app.OnMessageReactionRemoved(async (context, cancellationToken) =>
    {
        foreach (var reaction in context.Activity.ReactionsRemoved ?? [])
        {
            Console.WriteLine($"User removed reaction: {reaction.Type}");
        }
    });
    ```

If you only need a single handler that runs for both adds and removes, use `app.OnMessageReaction` instead.

A message sent from a reaction handler is placed relative to the reacted message.
In a 1:1 chat, when that message is inside a thread, the message is sent into the thread.
See [Responding to edits, reactions and card actions](../essentials/sending-messages.md#responding-to-edits-reactions-and-card-actions).

## Available Reaction Types

The SDK ships a small set of named reaction constants for the most common reactions, exposed via the `ReactionType` class.

Any string-valued reaction ID is accepted, so you can pass any reaction ID from the [Teams reactions reference page](https://learn.microsoft.com/en-us/microsoftteams/platform/agents-in-teams/teams-reactions-reference).

- `ReactionType.Like` — 👍
- `ReactionType.Heart` — ❤️
- `ReactionType.Checkmark` — ✅
- `ReactionType.Pushpin` — 📌
- `ReactionType.Laugh` — 😄
- `ReactionType.Surprise` — 😮
- `ReactionType.Sad` — 😢
- `ReactionType.Angry` — 😠
- `ReactionType.Hourglass` — ⏳
- `ReactionType.Exclamation` — ❗

## Skin Tones

Reactions tagged **Diverse** in the Teams reactions reference support five skin-tone variants. Append `-tone1` through `-tone5` to the reaction ID to select a variant:

```text
1f44b_wavinghand-tone4
```

## Rate Limits

Agents are limited to **two reactions per second**. Calls that exceed this limit are throttled and return a `429` response. Implement exponential backoff and honor the `Retry-After` header.

## Best Practices

- **Use reactions sparingly.** Reactions work best when they signal something specific (acknowledgement, completion, an error). Reacting to every message creates noise.
- **Match the reaction to the context.** Different reactions carry different meanings. Choose one that aligns with your agent's purpose.
- **Replace, don't stack.** If your agent has already added a reaction and needs to update it, remove the existing reaction first, then add the new one. Adding a reaction that's already present is a no-op.
- **Handle errors.** Reaction operations can fail if the message no longer exists or the agent isn't a member of the conversation. Handle exceptions appropriately.

## Differences from Feedback

Message reactions are different from the [feedback](./feedback.md) feature:

- **Reactions** are quick emoji responses your agent adds to or removes from messages programmatically, and reactions users add to messages your agent can listen for.
- **Feedback** are interactive UI components (like/dislike buttons) that users can click to provide structured feedback on agent responses.

Use reactions when your agent wants to acknowledge or respond to a message with emoji. Use feedback when you want to collect user opinions about your agent's responses.
