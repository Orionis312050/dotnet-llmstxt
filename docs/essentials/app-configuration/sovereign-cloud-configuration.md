---
title: "Sovereign Cloud Configuration"
sidebar_position: "1"
---

# Sovereign Cloud Configuration

:::info
Most developers do not need this page. It applies only to bots deployed in US Government (GCC-High, DoD) or China (21Vianet) cloud environments. If your bot runs in the standard commercial cloud, no additional configuration is needed.
:::

Sovereign clouds use separate Azure infrastructure with different service endpoints for authentication, token services, and bot communication. The Teams SDK handles this automatically when you specify your cloud environment.

## Supported Clouds

| Cloud | Value | Azure Portal | Teams Client |
|-------|-------|-------------|-------------|
| Public (default) | `Public` | portal.azure.com | teams.microsoft.com |
| US Gov (GCCH) | `USGov` | portal.azure.us | gov.teams.microsoft.us |
| US Gov (DoD) | `USGovDoD` | portal.azure.us | dod.teams.microsoft.us |
| China (21Vianet) | `China` | portal.azure.cn | n/a |

## Prerequisites

- A Teams tenant in the corresponding cloud environment
- An Azure Bot resource and app registration in the sovereign cloud portal

## Configuration

**SDK 2.0 (Legacy):**

The configuration below is only valid for **SDK 2.0** apps. It uses SDK 2.0 `Teams:Cloud` presets and per endpoint overrides. 
In `appsettings.json`:

```json
{
  "Teams": {
    "ClientId": "your-client-id",
    "ClientSecret": "your-client-secret",
    "TenantId": "your-tenant-id",
    "Cloud": "USGov",
    "LoginTenant": "your-tenant-id"
  }
}
```

Or programmatically:

```csharp
var app = new App(new AppOptions
{
    Cloud = CloudEnvironment.USGov,
    Credentials = new ClientCredentials("client-id", "client-secret")
});
```

**Available cloud presets:** `CloudEnvironment.Public`, `CloudEnvironment.USGov`, `CloudEnvironment.USGovDoD`, `CloudEnvironment.China`

Available override properties: `LoginEndpoint`, `LoginTenant`, `BotScope`, `TokenServiceUrl`, `OpenIdMetadataUrl`, `TokenIssuer`, `GraphScope`

**SDK 2.1 (current):**

For 2.1, configure sovereign cloud endpoints as shown below.

```json title="Properties/launchSettings.json (US Gov example)"
{
  "profiles": {
    "YourBot": {
      "commandName": "Project",
      "environmentVariables": {
        "AzureAd__Instance": "https://login.microsoftonline.us/",
        "AzureAd__TenantId": "your-tenant-id",
        "AzureAd__ClientId": "your-client-id",
        "BotFramework__OpenIdMetadataUrl": "https://login.botframework.azure.us/v1/.well-known/openid-configuration",
        "BotFramework__BotTokenIssuer": "https://api.botframework.us"
      }
    }
  }
}
```

For deployed environments, set equivalent values via app configuration or environment variables. When these aren't set, the SDK defaults to public-cloud endpoints.

## What the SDK Configures Automatically

When you set a cloud environment, the SDK automatically uses the correct endpoints for:

- **Login authority**: where tokens are acquired (e.g. `login.microsoftonline.us` for GCCH)
- **Bot token scope**: the OAuth scope for bot-to-service communication
- **Token service URL**: where user OAuth tokens are managed
- **JWT validation**: the signing keys and issuers used to verify inbound activity tokens
- **OpenID metadata**: the discovery endpoint for token validation configuration

You do not need to configure these endpoints individually.

## Next Steps

After configuring your cloud environment, your bot will authenticate and communicate using the correct sovereign cloud endpoints. No other code changes are needed; OAuth flows and all other SDK features work identically across clouds.

## Troubleshooting

### Bot fails to acquire a token at startup

**Symptom:** The SDK logs `AADSTS500011` (resource principal not found) or `AADSTS700016` (application not found in directory) when the app starts.

**Cause:** The bot's app registration lives in the wrong cloud's tenant. Sovereign clouds (USGov, USGovDoD, China) have separate Azure AD instances from the public cloud, and an app registered in the commercial cloud cannot authenticate to a sovereign tenant (and vice versa).

**Fix:** Register the bot in the sovereign cloud portal that matches your `cloud` setting. For USGov/DoD use `portal.azure.us`, for China use `portal.azure.cn`.

### Inbound activities are rejected with "issuer mismatch" or "audience invalid"

**Symptom:** Token validation fails on inbound `/api/messages` requests. Logs show `IDX10204` or `IDX10214`.

**Cause:** The bot is configured for the wrong cloud, so its JWT validator expects tokens signed by a different issuer than the one your sovereign Bot Channel Service uses.

**Fix:** Confirm the `cloud` setting matches the cloud your Azure Bot resource was created in. The SDK derives the expected issuer and JWKS URI from this setting; mismatching the cloud breaks token validation even when the credentials are otherwise valid.

### China bots fail user OAuth flows with the multi-tenant login URL

**Symptom:** OAuth or user-token flows fail in China with login redirects that do not resolve.

**Cause:** China sovereign cloud does not have a stable multi-tenant login alias. Single-tenant bots must point to their specific tenant explicitly via a `LoginTenant` override.

**Fix:** Override `loginTenant` on the China preset to your tenant ID.

For other authentication errors not specific to sovereign clouds, see the [App Authentication](../../teams/app-authentication/user-assigned-managed-identity-authentication-setup.md) guide.
