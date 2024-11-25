---
lab:
    title: 'Introduction'
    module: 'LAB 02: Integrate external content with Copilot for Microsoft 365 using Microsoft Graph connectors built with .NET'
---

# Introduction

Suppose you have an external system where you store knowledge base articles. These articles contain information about different processes in your organization. You want to be able to easily find and discover relevant information from Microsoft 365. You also want Copilot for Microsoft 365 to include information from these knowledge base articles in its responses.

To expose this external information inside Microsoft 365, you'll build a custom Microsoft Graph connector. Microsoft Graph connectors connect to your external system (1) to retrieve content, use the information from Microsoft Entra ID to authenticate with Microsoft 365 (2) and import the content to Microsoft 365 using the Microsoft Graph API (3).

![Diagram that shows conceptual working of a Microsoft Graph connector.](../media/1-graph-connector-concept.png)

In this module, you learn what Microsoft Graph connectors are and why you should consider using them in your organization. You build a Microsoft Graph connector that imports local markdown files to Microsoft 365. You also learn about how to ensure that the external content you import is accessible only to individuals with appropriate assigned permissions. Finally, you optimize your Microsoft Graph connector for use with Copilot for Microsoft 365.

## Prerequisites

- Basic knowledge of C#
- Basic knowledge of authentication
- Access to a [Microsoft 365 developer tenant](https://developer.microsoft.com/microsoft-365/dev-program?ocid=MSlearn)
- [.NET 8.0](https://dotnet.microsoft.com/download/dotnet/8.0)

When you're ready to begin, [continue to the next exercise...](./2-exercise-configure-connection-schema.md)