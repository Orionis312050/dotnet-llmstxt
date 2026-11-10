---
title: "Tabs"
sidebar_position: "6"
---

# Tabs

Tabs are host-aware webpages embedded in Microsoft Teams, Outlook, and Microsoft 365. Tabs are commonly implemented as Single Page Applications that use the Teams [JavaScript client library](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-library) (TeamsJS) to interact with the app host.

**SDK 2.0 (Legacy):**

SDK 2.0 includes a built-in `app.AddTab(...)` helper for hosting static tab content from your bot host.

**SDK 2.1 (current):**

SDK 2.1 does not expose a dedicated `AddTab` helper. Instead, host tab content with standard ASP.NET routes (for example `app.MapGet("/tabs/name/...")`) and keep using Teams SDK handlers for bot-side tab functions.

## Resources

- [Tabs overview](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/what-are-tabs?tabs=personal)
- [Teams JavaScript client library](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/using-teams-client-library)
- [Microsoft Graph overview](https://learn.microsoft.com/en-us/graph/overview)
- [Microsoft Authentication Library (MSAL)](https://learn.microsoft.com/en-us/entra/identity-platform/msal-overview)
- [Nested App Authentication (NAA)](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/authentication/nested-authentication)

### Additional resources

- [Static Pages](./server/static-pages.md)
- [TypeScript Tabs in-depth guide](./tabs.md)
