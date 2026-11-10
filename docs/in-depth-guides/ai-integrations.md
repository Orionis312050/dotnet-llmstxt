---
title: "AI Integrations"
sidebar_position: "1"
---

# AI Integrations

:::warning[Our AI libraries are deprecated]
The Teams SDK has deprecated its own AI libraries — the `@microsoft/teams.ai` packages (`ChatPrompt`, `Model`, and the older `@microsoft/teams.mcp` / `@microsoft/teams.a2a` plugins) — in favor of dedicated AI frameworks. Use the pattern shown in these guides instead: bring the OpenAI SDK (or any framework you like), and wire MCP and A2A directly into your Teams app.
:::

:::warning[SDK 2.0 AI libraries are deprecated]
For .NET, AI-specific libraries and older patterns from SDK 2.0 are deprecated. In SDK 2.1, use `Microsoft.Teams.Apps` as the Teams transport/runtime layer and integrate your preferred AI stack (for example Microsoft Agent Framework, OpenAI SDK, MCP, and A2A) in your handlers.
:::

Microsoft Teams SDK provides the platform and conversational interface for your app while remaining agnostic to the underlying intelligence. You can choose any AI framework, model, or protocol that suits your scenario and integrate it into your message handlers. The samples below walk through a few common ways to do that, from a single bot reasoning with an agent to multiple bots collaborating with each other.

- **[Build an agent in Teams](./ai-integrations/build-an-agent-in-teams.md)** — sample-backed by [`core/samples/ExtAIBot`](https://github.com/microsoft/teams.net/tree/main/core/samples/ExtAIBot), combining SDK 2.1 + Azure OpenAI + MCP tools.
- **[Enhancing the Teams Experience](./ai-integrations/enhance-the-teams-experience.md)** — same `ExtAIBot` sample shows streaming, AI-generated metadata, adaptive-card clarification, suggested follow-ups, citations, and feedback handling.
- **[Exposing Teams to AI Agents (MCP)](./ai-integrations/exposing-teams-to-ai-agents-mcp.md)** — sample-backed by [`core/samples/McpServer`](https://github.com/microsoft/teams.net/tree/main/core/samples/McpServer), using `AddMcpServer()` and `[McpServerTool]`.
- **[Bot-to-Bot with A2A](./ai-integrations/bot-to-bot-communication-with-a2a.md)** — sample-backed by [`core/samples/A2ABot`](https://github.com/microsoft/teams.net/tree/main/core/samples/A2ABot), wiring A2A + Teams proactive handoff.
