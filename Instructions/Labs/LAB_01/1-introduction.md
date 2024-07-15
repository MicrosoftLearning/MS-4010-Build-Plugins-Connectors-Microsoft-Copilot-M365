---
lab:
    title: 'Introduction'
    module: 'LAB 03: Connect Copilot for Microsoft 365 to your external data in real-time with message extension plugins built with .NET and Visual Studio'
---

# Introduction

Message extensions let users work with external systems from Microsoft Teams and Microsoft Outlook. Users can use message extensions to look up and change data and share the information from these systems in messages and emails as a rich formatted card.

Suppose you have a SharePoint Online List with product information that is current and relevant to your organization. You want to search and share this information across Microsoft 365. You also want Copilot for Microsoft 365 to use this information in its answers.

![Screenshot of the homepage from the Product support SharePoint Online team site. A list of recently released products is shown.](../media/1-sharepoint-online-product-support-site.png)

In this lab, you create a message extension. Your message extension uses a bot to communicate with Microsoft Teams, Microsoft Outlook, and Copilot for Microsoft 365.

![Screenshot of search results returned by a search based message extension in Microsoft Teams.](../media/2-search-results-nuget.png)

It uses Microsoft Entra to authenticate users, which enable it to return data from SharePoint Online using Microsoft Graph API on their behalf.

![Screenshot of an authentication challenge in a search based message extension. A link to sign-in is displayed.](../media/3-sign-in.png)

After the user authenticates, your message extension will get product information from SharePoint Online using Microsoft Graph API. It returns search results that can be embedded into messages and emails as a rich formatted card, and then shared.

![Screenshot of search results returned by a search based message extension in Microsoft Teams. The search results are returned from SharePoint Online. Each search result displays the product name, category, and product image.](../media/4-search-results-sharepoint-online.png)

![Screenshot of search result that is embedded in a message in Microsoft Teams. The search results are rendered as an Adaptive Card with the product name, category, call volume, and release date. An action button with the title View is displayed that users can use to navigate to the product list item in SharePoint Online.](../media/5-adaptive-card.png)

It works with Copilot for Microsoft 365 as a plugin, enabling it to query the SharePoint Online List on behalf of the user and use the returned data in its answers.

![Screenshot of an answer in Copilot for Microsoft 365 that contains information returned by the message extension plugin. An adaptive card is displayed showing product information.](../media/6-copilot-answer.png)

By the end of this lab, you're able to create message extensions written in C# (running on .NET). It can be used in Microsoft Teams, Microsoft Outlook, and Copilot for Microsoft 365. It can query data behind protected APIs and return the results as rich formatted cards.

## Prerequisites

- Basic knowledge of C#
- Basic knowledge of Bicep
- Basic knowledge of authentication
- Administrator access to a Microsoft 365 tenant
- Access to an Azure subscription
- Access to Copilot for Microsoft 365 is optional and required only to complete one exercise
- Visual Studio 2022 17.9 with [Teams Toolkit](/microsoftteams/platform/toolkit/toolkit-v4/teams-toolkit-fundamentals-vs) (Microsoft Teams development tools component) installed
- [.NET 8.0](https://dotnet.microsoft.com/download/dotnet/8.0)

When you're ready to begin, [continue to the next exercise...](./2-exercise-create-a-message-extension.md)
