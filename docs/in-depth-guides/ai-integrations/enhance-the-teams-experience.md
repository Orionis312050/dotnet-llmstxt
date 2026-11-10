---
title: "Enhance the Teams Experience"
sidebar_position: "2"
---

# Enhance the Teams Experience

:::warning[Our AI libraries are deprecated]
The Teams SDK has deprecated its own AI libraries — the `@microsoft/teams.ai` packages (`ChatPrompt`, `Model`, and the older `@microsoft/teams.mcp` / `@microsoft/teams.a2a` plugins) — in favor of dedicated AI frameworks. Use the pattern shown in these guides instead: bring the OpenAI SDK (or any framework you like), and wire MCP and A2A directly into your Teams app.
:::

:::warning[SDK 2.0 AI libraries are deprecated]
For .NET, implement AI enhancements with SDK 2.1 plus external model/agent frameworks. The Teams SDK provides the chat/activities surface; your AI stack provides reasoning/tool orchestration.
:::

You can enrich the agent output into a more Teams-native experience — adding structure, interactivity, and metadata on top of the generated text. This guide builds on the agent from [Build an agent in Teams](./build-an-agent-in-teams.md).

## Streaming

Streaming delivers responses to Teams incrementally as they're generated, rather than waiting for the full reply to complete. Each chunk of text is appended to the stream as it arrives.

```csharp
TeamsStreamingWriter writer = TeamsStreamingWriter.CreateFromContext(context);
await writer.SendInformativeUpdateAsync("Thinking…", cancellationToken);

await foreach (ChatResponseUpdate update in _chatClient.GetStreamingResponseAsync(history, options, cancellationToken))
{
    if (!string.IsNullOrEmpty(update.Text))
    {
        await writer.AppendResponseAsync(update.Text, cancellationToken);
    }
}
 await writer.FinalizeResponseAsync(cancellationToken: cancellationToken);
```

See [Streaming](../../essentials/sending-messages.md#streaming) for the full story on how Teams renders chunks and the constraints on stream lifecycle.

## AI-generated label

Mark the message as system-generated so Teams clearly labels it as AI output.

Mark the message as system-generated so Teams clearly labels it as AI output.

```csharp
MessageActivityInput reply = new MessageActivityInput().AddAIGenerated();
await writer.FinalizeResponseAsync(msg, cancellationToken);
```

<img
  src=StreamingImgUrl
  alt="Animated screenshot of an agent reply streaming into a Teams chat token by token, with the 'AI generated' label on the message."
  style={ width: '100%', maxWidth: 900 }
/>

## Feedback

Enable built-in thumbs up/down controls on the reply and surface a custom feedback form when users respond.

Enable built-in thumbs up/down controls on the reply and surface a custom feedback form when users respond.
```csharp
MessageActivityInput reply = new MessageActivityInput().AddAIGenerated().AddFeedback(FeedbackTypes.Custom);
await writer.FinalizeResponseAsync(msg, cancellationToken);
```

See [Feedback](../feedback.md) for the full form-handling story — capturing the submission, persisting it, and following up with the user.

## Clarification cards

When the agent calls the `request_clarification` tool (from [Build an agent](./build-an-agent-in-teams.md#adding-a-local-tool)), the reply is a card, not text. The model still produces a short wrap-up after the tool returns, so discard the streamed text and send only the card. Clearing the stream's accumulated text before emitting the card-only activity keeps the turn to a single clean reply.

```csharp
    private async Task RespondAsync(Context context, string userText, CancellationToken cancellationToken)
        where TActivity : TeamsActivity
    {
        _ = context.Activity.Conversation?.Id
            ?? throw new InvalidOperationException("Missing conversation ID.");

        TeamsStreamingWriter writer = TeamsStreamingWriter.CreateFromContext(context);
        RunResult result = await _agent.RunAsync(context.Activity.Conversation!.Id, userText, writer, cancellationToken);

        MessageActivityInput msg = new MessageActivityInput();

        if (result.PendingCards.Count > 0)
        {
            // Card-only reply (e.g. clarification). No text and no feedback — the card IS the question.
            msg.WithText("")
                .AddAttachment([.. result.PendingCards.Select(c =>
                    TeamsAttachment.CreateBuilder().WithAdaptiveCard(c).Build())])
                    .AddAIGenerated();
        }
        else
        {
            // normal reply: attach follow-ups, citations, feedback (below).
            ...
        }

        await writer.FinalizeResponseAsync(msg, cancellationToken);
    }
```

```csharp
this.OnAdaptiveCardAction(async (context, cancellationToken) =>
{
    if (context.Activity.Value?.Action?.Verb == "clarification")
    {
        string choice = context.Activity.Value.Action.Data?["clarificationChoice"]?.ToString() ?? "";
        await RespondAsync(context, choice, cancellationToken);
    }
    return InvokeResponse.Ok();
});
```

The user's selection arrives as a fresh turn through the card-action route — the same code path as a normal message — so the agent picks up with full context.

<img
  src=ClarificationImgUrl
  alt="Animated screenshot of the clarification flow: the user asks an ambiguous question, the bot replies with a choice card, the user picks an option, and the bot streams a grounded answer with an inline citation."
  style={ width: '100%', maxWidth: 900 }
/>

## Suggested prompts

Suggested prompts give the user one-click follow-up questions after a reply. In Teams they render as chips under the message; tapping one sends the value back as a normal user message, so the same message handler picks it up — no extra routing required.

Rather than hard-coding them, generate two contextual follow-ups with a separate lightweight model call constrained to a strict JSON schema, then attach them as suggested actions.

```csharp
private const string FollowUpsPrompt = """
    Produce 2 specific prompts the user might want to ask next.

    Output format — read carefully:
    Return ONLY a JSON object INSTANCE, like this:
    {"prompt1": "How do I stream a reply?", "prompt2": "Show me an Adaptive Card example"}

    Each prompt MUST:
    - Be phrased in the first person, as the user would type.
    - Stay under 8 words.

    Pick based on the conversation:
    - If recent turns have substantive content, drill into a concrete topic, API, or
    concept that just came up.
    - Otherwise (e.g. conversation just started, or the last turn is generic),
    suggest prompts that showcase what you can help with based on the MCP tools available.
    """;

private async Task> GenerateFollowUpsAsync(
    IReadOnlyList history,
    CancellationToken cancellationToken)
{
    List messages =
    [
        .. history,
        new ChatMessage(ChatRole.System, FollowUpsPrompt)
    ];

    ChatResponse response = await _chatClient.GetResponseAsync(
        messages,
        cancellationToken: cancellationToken);

    if (!response.TryGetResult(out FollowUps? followUps) || followUps is null)
    {
        _logger.LogWarning("Follow-up generation did not return parseable JSON. Raw response: {Text}", response.Text);
        return [];
    }

    return [
        new SuggestedAction(ActionTypes.IMBack, followUps.Prompt1),
        new SuggestedAction(ActionTypes.IMBack, followUps.Prompt2)
    ];
}
```

<img
  src=SuggestedPromptsImgUrl
  alt="Animated screenshot of suggested follow-up prompt chips appearing under an agent reply; tapping one sends it back as the next user message."
  style={ width: '100%', maxWidth: 900 }
/>

## Citations

Citations render as footnote-style references inline with the reply — `[1]`, `[2]`, etc. — surfacing the source title, abstract, and URL on hover. They originate from tool outputs, where the collector from [Grounding responses with citations](./build-an-agent-in-teams.md#grounding-responses-with-citations) assigned each result a stable position.

When building the final reply, attach only the citations whose position actually appears in the streamed text.

```csharp
result.Citations.AttachCitations(reply, result.FullText);

public void AttachCitations(MessageActivityInput reply, string fullText)
{
    HashSet<int> used = [];
    foreach (Match match in Regex.Matches(fullText, @"\[(\d+)\]"))
    {
        if (int.TryParse(match.Groups[1].Value, out int position))
            used.Add(position);
    }

    foreach (CitationEntry citation in _citations.Values.Where(e => used.Contains(e.Position)))
    {
        reply.AddCitation(
            citation.Position,
            new CitationAppearance
            {
                Name = string.IsNullOrEmpty(citation.Title)
                    ? $"Source {citation.Position}"
                    : citation.Title[..Math.Min(80, citation.Title.Length)],
                Abstract = string.IsNullOrEmpty(citation.Snippet)
                    ? "No description available."
                    : citation.Snippet,
                Url = Uri.TryCreate(citation.Url, UriKind.Absolute, out Uri? uri) ? uri : null
            });
    }
}
```

<img
  src=CitationsImgUrl
  alt="Animated screenshot showing a user hovering over a footnote citation in an agent response, with a pop-up showing explanatory text."
  style={ width: '100%', maxWidth: 900 }
/>

## Sensitivity labels

:::info[Coming Soon]
Sensitivity labels are planned for the next SDK release, version 2.2.0.
:::

Sensitivity labels mark a message with a Microsoft Purview label, which Teams shows as a shield on the message with the label name on hover.

Add a message-level label with `AddSensitivityLabel`, passing the label ID as `Id`, which is sent to Teams as `@id`. Calling it again replaces the previous label. To label an individual source, set `SensitivityUsageInfo` on the citation appearance.

```csharp
string labelId = "<purview-label-id>";

MessageActivityInput reply = new MessageActivityInput()
    .WithText("Here is the quarterly summary [1].")
    .AddSensitivityLabel(new SensitivityUsageInfo { Id = labelId, Name = "Confidential" })
    .AddCitation(1, new CitationAppearance
    {
        Name = "Quarterly report",
        Abstract = "Revenue and forecast for the quarter.",
        SensitivityUsageInfo = new SensitivityUsageInfo { Id = labelId, Name = "Confidential" }
    });
```

- `Name` is optional but recommended. Teams shows it when it cannot resolve `@id`, for example for a deleted label, and shows no label at all if there is no `Name` either. Set it whenever you set `Description`, because Teams can reject a message whose label has a description but no name.
- The sensitivity label, the AI-generated label, and citations are all written to the same root message entity, so an activity always carries exactly one root message entity, regardless of the order you call them in.

### Where label IDs come from

The label ID is the GUID of a Purview sensitivity label configured in your tenant. Get it from the Microsoft Purview portal or from your tenant admin.
