---
title: "Agentic User lifecycle events"
sidebar_position: "4"
---

# Agentic User lifecycle events

Agent 365 sends lifecycle events when an Agentic User is created or when its identity, manager, availability, deletion status, or workload onboarding changes. These events let your app keep its state and behavior synchronized with Microsoft 365.

For example, your app can initialize per-agent storage after creation, update ownership when the manager changes, pause scheduled work when the Agentic User is disabled, clean up access after deletion, and wait for Teams onboarding before sending messages.

## Supported events

| Event                         | Typical use                                           |
| ----------------------------- | ----------------------------------------------------- |
| Agentic User identity created | Initialize per-agent configuration and state          |
| Identity property updated     | Refresh mail, alias, or user principal name data      |
| Manager updated               | Update ownership, approval, or escalation routing     |
| Enabled                       | Resume work for the Agentic User                      |
| Disabled                      | Stop scheduled or proactive work                      |
| Deleted                       | Disable access and begin cleanup                      |
| Undeleted                     | Restore state after the Agentic User is restored      |
| Workload onboarding updated   | Wait for Teams onboarding to succeed before messaging |

The event names in activity payloads use `AgenticUser*`, such as `AgenticUserIdentityCreated` and `AgenticUserDisabled`.

## Register handlers

Register a general handler when you need common processing for every lifecycle event. Register typed handlers when you need the payload for a specific variant.

```csharp
teamsApp.OnAgentLifecycle((context, _) =>
{
    context.Log.LogInformation(
        "Lifecycle event: {ValueType}",
        context.Activity.ValueType);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserIdentityCreated((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User created: {AgenticUserId}",
        context.Activity.Value?.AgenticUserId);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserIdentityUpdated((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User identity updated: {UpdatedProperty}",
        context.Activity.Value?.UpdatedProperty);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserManagerUpdated((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User manager updated: {Manager}",
        context.Activity.Value?.Manager);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserEnabled((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User enabled: {AgenticUserId}",
        context.Activity.Value?.AgenticUserId);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserDisabled((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User disabled: {AgenticUserId}",
        context.Activity.Value?.AgenticUserId);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserDeleted((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User deleted: {DeletionReason}",
        context.Activity.Value?.DeletionReason);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserUndeleted((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User restored: {AgenticUserId}",
        context.Activity.Value?.AgenticUserId);
    return Task.CompletedTask;
});

teamsApp.OnAgenticUserWorkloadOnboardingUpdated((context, _) =>
{
    context.Log.LogInformation(
        "Agentic User workload onboarding updated: {WorkloadName} ({State})",
        context.Activity.Value?.WorkloadName,
        context.Activity.Value?.WorkloadOnboardingState);
    return Task.CompletedTask;
});
```
