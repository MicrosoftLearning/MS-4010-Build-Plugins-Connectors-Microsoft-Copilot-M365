---
lab:
    title: 'Exercise 3 - Ensure secure access with access control list'
    module: 'LAB 02: Integrate external content with Copilot for Microsoft 365 using Microsoft Graph connectors built with .NET'
---

# Exercise 3 - Ensure secure access with access control list

In this exercise, you update the code responsible for importing local markdowns files with configuring ACLs on selected items.

## Before you start

This exercise will take you about **XX minutes** to complete.

## Task 1 - Import content that's available to everyone in the organization

When you implemented the code for importing external content in the previous exercise, you configured it to be available to everyone in the organization. Here's the code that you used:

```csharp
static IEnumerable<ExternalItem> Transform(IEnumerable<DocsArticle> content)
{
  var baseUrl = new Uri("https://learn.microsoft.com/graph/");

  return content.Select(a =>
  {
    var docId = GetDocId(a.RelativePath ?? "");

    return new ExternalItem
    {
      Id = docId,
      Properties = new()
      {
        AdditionalData = new Dictionary<string, object> {
            { "title", a.Title ?? "" },
            { "description", a.Description ?? "" },
            { "url", new Uri(baseUrl, a.RelativePath!.Replace(".md", "")).ToString() }
        }
      },
      Content = new()
      {
        Value = a.Content ?? "",
        Type = ExternalItemContentType.Html
      },
      Acl = new()
      {
          new()
          {
            Type = AclType.Everyone,
            Value = "everyone",
            AccessType = AccessType.Grant
          }
      }
    };
  });
}
```

You configured the ACL to grant access to everyone. Let's adjust it for selected markdown pages that you're importing.

## Task 2 - Import content available to select users

First, configure one of the pages that you're importing to be accessible only to a specific user.

In the web browser:

1. Navigate to the Azure portal at [https://portal.azure.com](https://portal.azure.com) and sign in with your work or school account.
1. From the sidebar, select **Microsoft Entra ID**.
1. From the navigation, select **Users**.
1. From the list of users, open one of the users by selecting their name.
1. Copy the value of the **Object ID** property.

   :::image type="content" source="../media/8-user.png" alt-text="Screenshot of the Azure portal with a user profile open.":::

Use this value to define a new ACL for a specific markdown page.

In the code editor:

1. Open the **ContentService.cs** file, and locate the `Transform` method.
1. Inside the `Select` delegate, define the default ACL that applies to all imported items:

   ```csharp
   var acl = new Acl
   {
     Type = AclType.Everyone,
     Value = "everyone",
     AccessType = AccessType.Grant
   };
   ```

1. Next, override the default ACL for the markdown file with name ending in `use-the-api.md`:

   ```csharp
   if (a.RelativePath!.EndsWith("use-the-api.md"))
   {
     acl = new()
     {
       Type = AclType.User,
       // AdeleV
       Value = "6de8ec04-6376-4939-ab47-83a2c85ab5f5",
       AccessType = AccessType.Grant
     };
   }
   ```

1. Finally, update the code that returns the external item to use the defined ACL:

   ```csharp
   return new ExternalItem
   {
     Id = docId,
     Properties = new()
     {
       AdditionalData = new Dictionary<string, object> {
         { "title", a.Title ?? "" },
         { "description", a.Description ?? "" },
         { "url", new Uri(baseUrl, a.RelativePath!.Replace(".md", "")).   ToString() }
       }
     },
     Content = new()
     {
       Value = a.Content ?? "",
       Type = ExternalItemContentType.Html
     },
     Acl = new()
     {
       acl
     }
   };
   ```

1. The updated `Transform` method looks as follows:

   ```csharp
   static IEnumerable<ExternalItem> Transform(IEnumerable<DocsArticle>    content)
   {
     var baseUrl = new Uri("https://learn.microsoft.com/graph/");
   
     return content.Select(a =>
     {
       var acl = new Acl
       {
         Type = AclType.Everyone,
         Value = "everyone",
         AccessType = AccessType.Grant
       };
   
       if (a.RelativePath!.EndsWith("use-the-api.md"))
       {
         acl = new()
         {
           Type = AclType.User,
           // AdeleV
           Value = "6de8ec04-6376-4939-ab47-83a2c85ab5f5",
           AccessType = AccessType.Grant
         };
       }
   
       var docId = GetDocId(a.RelativePath ?? "");
   
       return new ExternalItem
       {
         Id = docId,
         Properties = new()
         {
           AdditionalData = new Dictionary<string, object> {
             { "title", a.Title ?? "" },
             { "description", a.Description ?? "" },
             { "url", new Uri(baseUrl, a.RelativePath!.Replace(".md", "")).   ToString() }
           }
         },
         Content = new()
         {
           Value = a.Content ?? "",
           Type = ExternalItemContentType.Html
         },
         Acl = new()
         {
           acl
         }
       };
     });
   }
   ```

1. Save your changes.

## Task 3 - Import content available to a select group

Now, let's extend the code so that another page is only accessible by a select group of users.

In the web browser:

1. Navigate to the Azure portal at [https://portal.azure.com](https://portal.azure.com) and sign in with your work or school account.
1. From the sidebar, select **Microsoft Entra ID**.
1. From the navigation, select **Groups**.
1. From the list of groups, open one of the groups by selecting their name.
1. Copy the value of the **Object Id** property.

:::image type="content" source="../media/8-group.png" alt-text="Screenshot of the Azure portal with a group page open.":::

Use this value to define a new ACL for a specific markdown page.

In the code editor:

1. Open the **ContentService.cs** file, and locate the `Transform` method
1. Extend the previously defined `if` clause, with an extra condition to define the ACL for the markdown file with name ending in `traverse-the-graph.md`:

   ```csharp
   if (a.RelativePath!.EndsWith("use-the-api.md"))
   {
     acl = new()
     {
       Type = AclType.User,
       // AdeleV
       Value = "6de8ec04-6376-4939-ab47-83a2c85ab5f5",
       AccessType = AccessType.Grant
     };
   }
   else if (a.RelativePath.EndsWith("traverse-the-graph.md"))
   {
     acl = new()
     {
       Type = AclType.Group,
       // Sales and marketing
       Value = "a9fd282f-4634-4cba-9dd4-631a2ee83cd3",
       AccessType = AccessType.Grant
     };
   }
   ```

1. The updated `Transform` method looks as follows:

   ```csharp
   static IEnumerable<ExternalItem> Transform(IEnumerable<DocsArticle>    content)
   {
     var baseUrl = new Uri("https://learn.microsoft.com/graph/");
   
     return content.Select(a =>
     {
       var acl = new Acl
       {
         Type = AclType.Everyone,
         Value = "everyone",
         AccessType = AccessType.Grant
       };
   
       if (a.RelativePath!.EndsWith("use-the-api.md"))
       {
         acl = new()
         {
           Type = AclType.User,
           // AdeleV
           Value = "6de8ec04-6376-4939-ab47-83a2c85ab5f5",
           AccessType = AccessType.Grant
         };
       }
       else if (a.RelativePath.EndsWith("traverse-the-graph.md"))
       {
         acl = new()
         {
           Type = AclType.Group,
           // Sales and marketing
           Value = "a9fd282f-4634-4cba-9dd4-631a2ee83cd3",
           AccessType = AccessType.Grant
         };
       }
   
       var docId = GetDocId(a.RelativePath ?? "");
   
       return new ExternalItem
       {
         Id = docId,
         Properties = new()
         {
           AdditionalData = new Dictionary<string, object> {
               { "title", a.Title ?? "" },
               { "description", a.Description ?? "" },
               { "url", new Uri(baseUrl, a.RelativePath!.Replace(".md",    "")).ToString() }
           }
         },
         Content = new()
         {
           Value = a.Content ?? "",
           Type = ExternalItemContentType.Html
         },
         Acl = new()
         {
             acl
         }
       };
     });
   }
   ```

1. Save your changes.

## Task 4 - Apply the new ACLs

The final step is to apply the newly configured ACLs.

1. Open a terminal and change the working directory to your project.
1. Build the project by running the `dotnet build` command.
1. Start loading the content by running the `dotnet run -- load-content` command.

[Continue to the next exercise...](./5-exercise-enable-inline-results.md)