---
title: "Sending Inline Images"
sidebar_position: "4"
---

# Sending Inline Images

Your agent can send an image that renders directly in the conversation, beside the message text, instead of arriving as a document the user has to open. Teams calls this an inline image. It is the outbound counterpart to [Receiving Files](./receiving-files.md): receiving gives you a lazy handle you pull bytes from, while sending means putting the image on the message before it goes out.

An inline image travels as an attachment on the outgoing message activity. Two fields do the work: `ContentType` declares an image MIME type, which is what tells Teams to render the attachment as a picture rather than list it as a file, and `ContentUrl` carries the image itself in one of two forms:

- **A hosted URL**, a reachable `https://` link the client fetches. Use this whenever the image already lives somewhere, and for anything that is not small.
- **A base64 data URI**, the bytes embedded in the message as `data:image/png;base64,...`. Use this when the bytes only exist inside your process, such as a chart you just rendered, and you have nowhere to host them.

Sending an inline image does not upload anything to OneDrive or SharePoint. The image is part of the message rather than a stored document, so it does not appear in the conversation's **Files** tab.

:::info[Preview]
Unlike [receiving files](./receiving-files.md), which a traditional bot can do only in personal (1:1) chats, sending is not scope-restricted: the same call works in personal chats, group chats, and channels.
:::

Sending an image needs no Microsoft Graph permission and no user sign-in. The message goes out under your app's own identity and is authorized by your app being in the conversation, exactly like sending text.

## Image limits

Teams renders inline images within fixed platform limits. Images can be up to **1024x1024 pixels** and **1 MB**, in **PNG**, **JPEG**, or **GIF** format. Animated GIFs are not supported. No validation is performed on MIME type, dimensions, or size of an outgoing attachment, so an image outside these bounds may send successfully but fail to render on the client.

:::warning
Base64 encoding inflates the bytes by roughly a third, and unlike a hosted URL the encoded string travels inside the message payload itself rather than being fetched separately. The practical ceiling for an inline data URI is therefore well below the 1 MB image limit. Keep embedded images small (icons, thumbnails, simple charts) and prefer a hosted URL for anything larger.
:::

## Send an image from a hosted URL

Build an attachment whose `ContentUrl` is a link the Teams client can reach, and add it to the message you send:

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Schema;

teamsApp.OnMessage(async (context, cancellationToken) =>
{
    TeamsAttachment image = TeamsAttachment.CreateBuilder()
        .WithContentType(new AttachmentContentType("image/png"))
        .WithContentUrl(new Uri("https://contoso.com/charts/weekly.png"))
        .WithName("weekly.png")
        .Build();

    await context.SendAsync(
        new MessageActivityInput()
            .WithText("Here is the latest chart:")
            .AddAttachment(image),
        cancellationToken);
});
```

`AttachmentContentTypes` exposes constants for the card content types only, so an image type is constructed directly: `new AttachmentContentType("image/png")`.

The URL must be reachable by the client without your agent's credentials. A link behind a login, or one that only resolves inside your network, renders as a broken image.

## Send image bytes inline

When the bytes only exist in memory or on your agent's disk, encode them as base64 and send them as a data URI. The MIME type appears twice: once in the attachment's content type, and once inside the data URI. Keep them consistent with each other and with the actual bytes:

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Apps.Schema;

teamsApp.OnMessage(async (context, cancellationToken) =>
{
    byte[] bytes = await File.ReadAllBytesAsync("./charts/weekly.png", cancellationToken);
    string dataUri = $"data:image/png;base64,{Convert.ToBase64String(bytes)}";

    TeamsAttachment image = TeamsAttachment.CreateBuilder()
        .WithContentType(new AttachmentContentType("image/png"))
        .WithContentUrl(new Uri(dataUri))
        .WithName("weekly.png")
        .Build();

    await context.SendAsync(
        new MessageActivityInput()
            .WithText("Here is the latest chart:")
            .AddAttachment(image),
        cancellationToken);
});
```

:::note
`ContentUrl` is a `Uri`, not a `string`, so a data URI has to be wrapped in `new Uri(...)`. `data:` is a scheme `Uri` understands, and the base64 payload survives the round trip unchanged, including its `+`, `/`, and `=` characters.
:::

## Position an image inside the text

The paths above place the image alongside the message rather than at a chosen point in it. To control placement, write the message body as HTML and set the text format to `xml`. Each image renders exactly where its `<img>` tag sits:

```csharp
using Microsoft.Teams.Apps;

teamsApp.OnMessage(async (context, cancellationToken) =>
{
    string encoded = Convert.ToBase64String(
        await File.ReadAllBytesAsync("charts/weekly.png", cancellationToken));

    await context.SendAsync(
        new MessageActivityInput()
            .WithText($"<div>Revenue is up.<img src=\"data:image/png;base64,{encoded}\"/>Questions?</div>")
            .WithTextFormat(TextFormats.Xml),
        cancellationToken);
});
```

The `xml` form also lets you size each image, with `height` and `width` attributes on the tag. A Markdown image gives you neither placement nor sizing and renders at a default 256x256, so `<img>` is the only form that controls either.

Two constraints apply to this path only:

- **Base64 only.** Only `<img>` tags whose `src` is a base64 data URI are uploaded and rewritten for you. An `https://` source in `xml` text is not, so use the hosted-URL attachment above for images you host yourself.
- **The payload ceiling still applies**, and more sharply: every embedded image counts against the same message budget as your text. Exceeding the platform's per-message image limit rejects the **whole message** rather than dropping the extra images.

:::warning[Restricted clouds]
Agents do not support sending and receiving **files** in GCC High, DoD, and Teams operated by 21Vianet. In those environments a base64 image attachment is the only way to get an image into a message. See [Send and receive files](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/bots-filesv4).
:::

## Attachment fields

An image attachment uses a small subset of the attachment shape:

| Property | Description |
|---|---|
| `ContentType` | Required. The image's MIME type (`image/png`, `image/jpeg`, or `image/gif`), as an `AttachmentContentType`. This is what makes Teams render the attachment as a picture, so it has to match the actual bytes. |
| `ContentUrl` | The image itself, as a `Uri`: either a reachable `https://` URL, or a `data:<mime>;base64,<encoded>` URI. |
| `Name` | Optional display name for the attachment (e.g. `weekly.png`). |
| `ThumbnailUrl` | Optional preview image. Rarely needed for an inline image, since the image is already its own preview. |
| `Content` | Embedded payload, used by card attachments. Leave it unset for an image. |

## Send several images

Add more than one attachment to send a set of images in a single message, and set the attachment layout to control how the client arranges them:

```csharp
TeamsAttachment salesChart = TeamsAttachment.CreateBuilder()
    .WithContentType(new AttachmentContentType("image/png"))
    .WithContentUrl(new Uri("https://contoso.com/charts/sales.png"))
    .WithName("sales.png")
    .Build();

TeamsAttachment trafficChart = TeamsAttachment.CreateBuilder()
    .WithContentType(new AttachmentContentType("image/png"))
    .WithContentUrl(new Uri("https://contoso.com/charts/traffic.png"))
    .WithName("traffic.png")
    .Build();

await context.SendAsync(
    new MessageActivityInput()
        .WithText("This week at a glance:")
        .WithAttachmentLayout(AttachmentLayoutType.Carousel)
        .AddAttachment(salesChart, trafficChart),
    cancellationToken);
```

`AddAttachment()` takes `params TeamsAttachment[]` and appends, so calling it more than once builds the list up rather than replacing it. `AttachmentLayoutType` offers `List` (the default), `Carousel`, and `Grid`.

## Images in an Adaptive Card

A message attachment is not the only way to put a picture in front of a user. An [Adaptive Card](../adaptive-cards.md) can contain an `Image` element, which is a different mechanism with a different trade-off:

| | Message attachment | Adaptive Card image |
|---|---|---|
| Where it renders | Beside the message text, as a standalone picture | Inside the card's layout, alongside other card elements |
| What you control | Little. The client decides presentation. | Size, alignment, alt text, expand behavior, and a select action |
| Good for | A chart or screenshot that is the whole point of the message | An image that is one part of a richer, interactive message |

```csharp
using Microsoft.Teams.Apps;
using Microsoft.Teams.Cards;

AdaptiveCard card = new AdaptiveCard
{
    Body =
    [
        new Image("https://contoso.com/charts/weekly.png")
        {
            AltText = "Weekly sales chart",
            Size = Size.Large,
        }
    ],
};

await context.SendAsync(
    new MessageActivityInput().AddAdaptiveCardAttachment(card),
    cancellationToken);
```

Card images accept the same two `url` forms, a hosted URL or a base64 data URI, and are subject to the same practical payload ceiling when embedded.

## Streamed messages

Attachments ride on the final message of a stream, not on the intermediate chunks: the chunks are typing activities, which carry no attachments. Build up the text with the usual streaming calls, then emit a message activity with the image attached as the last one, and the image arrives with the completed message.

## Next steps

- [Receiving Files](./receiving-files.md) covers the inbound direction: reading files a user attaches to a message.
- [Sending Messages](../../essentials/sending-messages.md) covers the rest of the outbound message surface, including streaming, mentions, and quoted replies.
- [Adaptive Cards](../adaptive-cards.md) covers building richer card layouts.
- [Send and receive messages](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/conversations/conversation-messages#use-picture-messages) documents the platform-side picture limits.
- [Add media attachments to messages](https://learn.microsoft.com/en-us/azure/bot-service/rest-api/bot-framework-rest-connector-add-media-attachments?view=azure-bot-service-4.0) documents the underlying attachment wire format.
