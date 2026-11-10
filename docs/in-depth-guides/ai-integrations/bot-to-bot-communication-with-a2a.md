---
title: "Bot-to-Bot Communication with A2A"
sidebar_position: "4"
---

# Bot-to-Bot Communication with A2A

:::warning[Our AI libraries are deprecated]
The Teams SDK has deprecated its own AI libraries — the `@microsoft/teams.ai` packages (`ChatPrompt`, `Model`, and the older `@microsoft/teams.mcp` / `@microsoft/teams.a2a` plugins) — in favor of dedicated AI frameworks. Use the pattern shown in these guides instead: bring the OpenAI SDK (or any framework you like), and wire MCP and A2A directly into your Teams app.
:::

:::warning[SDK 2.0 AI libraries are deprecated]
For .NET, implement A2A with SDK 2.1 plus an external A2A protocol library. Use Teams SDK handlers for inbound user turns and proactive messaging, and use A2A messages for bot-to-bot handoff coordination.
:::

Agents are typically designed to interact either with people (chatbots) or with systems (tools, APIs, MCP servers). [Agent2Agent](https://a2a-protocol.org/) (A2A) introduces a third interaction model: agents communicating directly with other agents as peers — each with its own model, capabilities, and human audience.

This guide walks through a **handoff** between two Teams bots, **Alice** and **Bob**, each backed by its own LLM agent. A user DMs one bot; its agent reads the peer's capability description and decides whether to answer directly or hand the user off. On handoff, the receiving bot **proactively opens a 1:1 chat** with the user and greets them with the context that came across — so the conversation continues seamlessly in the new chat.

This guide is based on [`core/samples/A2ABot`](https://github.com/microsoft/teams.net/tree/main/core/samples/A2ABot): two SDK 2.1 bots run the same code with different config and hand users off through A2A.

## Advertising capabilities with an Agent Card

Every A2A server publishes an `AgentCard` — a small machine-readable document describing who the agent is and what it can do. Peers fetch this card to learn about each other; their LLMs then read the `description` field to decide *when* to hand off a user.

The sample publishes an A2A `AgentCard` per bot:

```csharp
public static AgentCard Build(Config config) => new()
{
    Name = config.Name,
    Description = config.Description,
    SupportedInterfaces =
    [
        new AgentInterface
        {
            Url = $"{config.SelfUrl}/a2a",
            ProtocolBinding = "JSONRPC",
            ProtocolVersion = "1.0",
        }
    ],
    Skills =
    [
        new AgentSkill
        {
            Id = "handoff",
            Name = "Handoff",
            Description = $"Accepts handoffs of users from peer bots. Specialty: {config.Description}",
        }
    ],
};
```

The `description` is the most important knob in this sample — it's the natural-language summary another bot's LLM uses to decide whether *this* bot is the right peer for a given user. Tweak it to match the persona and expertise you want each bot to advertise.

## The handoff message contract

A handoff carries everything the receiving bot needs to reach the user proactively: their **`aadObjectId`** (the tenant-wide identity both bots share — the Teams MRI one bot sees isn't valid against the other), the **`tenantId`**, the **`serviceUrl`**, and a **`summary`** of the conversation so the peer can pick up cold.

Use a strict handoff payload contract:

```csharp
internal record HandoffMessage(
    string Kind,
    string AadObjectId,
    string UserName,
    string Summary,
    string From,
    string TenantId,
    string ServiceUrl);
```

## LLM-driven handoff

Routing is not a hard-coded rule — the LLM decides. Each bot exposes a single `handoff_to_peer` tool to its agent, and the agent's instructions include the live `AgentCard.description` of the peer. When a question fits the peer's expertise better than its own, the model calls the tool.

The model gets one handoff tool that calls A2A:

```csharp
AIFunction handoffTool = AIFunctionFactory.Create(HandoffToPeerAsync, new AIFunctionFactoryOptions
{
    Name = "handoff_to_peer",
    Description = $"Hand off the current user to {_config.PeerName} when {_config.PeerName}'s expertise is a better fit.",
});
```

Then send structured handoff data:

```csharp
await _a2aClient.SendHandoffAsync(
    new HandoffMessage("handoff", turn.AadObjectId, turn.UserName, summary, _config.Name, turn.TenantId, turn.ServiceUrl),
    ct);
```

The identity needed to build the handoff is captured from the inbound Teams activity and stashed for the duration of the turn, so the tool callback can reach it without threading it through every call. A handoff *greeting* runs with no identity set — the tool guards against that to prevent a ping-pong.

## Sending a handoff over A2A

The outbound side resolves the peer's `AgentCard` once (so the agent can read its live description into the tool), then ships the handoff as a `DataPart`.

The outbound client resolves and caches the peer card once:

```csharp
A2ACardResolver resolver = new(new Uri(config.PeerUrl), http);
AgentCard card = await resolver.GetAgentCardAsync(ct);
global::A2A.A2AClient client = new(new Uri(card.SupportedInterfaces[0].Url), http);
```

Then posts the handoff as a `Data` part.

## Receiving a handoff

The inbound side implements the A2A protocol's executor interface. For each inbound message it pulls the handoff out of the `DataPart`, opens a fresh 1:1 with the user against **their** `serviceUrl`, asks the agent to seed that conversation's history with the handoff context and produce a greeting, sends the greeting proactively, and acks back so the sender's call resolves.

Inbound A2A creates a 1:1 conversation and sends a proactive greeting:

```csharp
CreateConversationResponse conv = await conversations.CreateConversationAsync(
    new ConversationParameters
    {
        IsGroup = false,
        TenantId = handoff.TenantId,
        Members = [new TeamsChannelAccount { Id = handoff.AadObjectId }],
    },
    serviceUrl,
    cancellationToken: ct);

string greeting = await agent.GreetWithHandoffAsync(newConvId, handoff.From, handoff.UserName, handoff.Summary, ct);
await conversations.SendActivityAsync(newConvId, new MessageActivityInput().WithText(greeting), serviceUrl, cancellationToken: ct);
```

Because the greeting turn is left in the per-conversation history, when the user replies in their new DM the agent picks up coherently.

## Wiring A2A into your Teams app

The Teams bot and A2A server run in the same process and share one HTTP surface: `/api/messages` for Teams, `/a2a` for inbound handoffs, and `/.well-known/agent-card.json` for the AgentCard.

Teams + A2A are hosted in one ASP.NET app:

```csharp
builder.Services.AddTeamsBotApplication();
builder.Services.AddA2AAgent(agentCard);

WebApplication webApp = builder.Build();
TeamsBotApplication teamsApp = webApp.UseTeamsBotApplication();

webApp.MapA2A("/a2a");
webApp.MapWellKnownAgentCard(agentCard);
webApp.Run();
```

Each bot needs its own Teams app registration (so DMs route to the right bot) and its own port. The sample runs Alice on `3978` and Bob on `3979`; their peer URLs point at each other.

## Putting it all together

With both bots running and installed for the user, DM Alice with a question outside her specialty and watch the round-trip: Alice's LLM calls `handoff_to_peer`, Bob receives it over A2A, opens a new 1:1 with the user, and greets them with an answer already in hand. The bots are symmetric — the same flow runs the other way from Bob to Alice.

<img
  src=A2AImgUrl
  alt="Animated screenshot of the end-to-end A2A handoff: a user DMs Alice, Alice hands off to Bob, and Bob opens a new chat greeting the user with context."
  style={ width: '100%', maxWidth: 900 }
/>

:::warning
This sample configures no authenticator on the A2A endpoint, so any caller can post a handoff. For production, validate the caller's identity (a bearer token or mTLS) before opening a conversation with someone they named.
:::
