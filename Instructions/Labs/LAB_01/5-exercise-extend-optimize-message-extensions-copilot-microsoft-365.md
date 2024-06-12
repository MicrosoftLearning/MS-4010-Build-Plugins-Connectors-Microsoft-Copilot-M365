---
lab:
    title: 'Exercise 4 - Extend and optimize message extensions for use with Copilot for Microsoft 365'
    module: 'LAB 03: Connect Copilot for Microsoft 365 to your external data in real-time with message extension plugins built with .NET and Visual Studio'
---

# Exercise 4 - Extend and optimize message extensions for use with Copilot for Microsoft 365

In this exercise, you extend and optimize your message extension for use with Copilot for Microsoft 365. You add a new parameter called Target Audience and update the message extension logic to handle multiple parameters. Finally, you run and debug your message extension and test it in Copilot in Microsoft Teams.

## Task 1 - Update app manifest

Specifying concise and accurate descriptions in your app manifest is critical to ensuring Copilot knows when and how to invoke your plugin. Update the app, command, and parameter descriptions in the app manifest.

Open Visual Studio:

1. In the **appPackage** folder, open the file named **manifest.json**
1. Update the **description** object

    ```json
    {
        "description": {
            "short": "Product look up tool.",
            "full": "Get real-time product information and share them in a conversation. Search by product name or target audience. ${{APP_DISPLAY_NAME}} works with Microsoft 365 Chat. Find products at Contoso. Find Contoso products called mark8. Find Contoso products named mark8. Find Contoso products related to Mark8. Find Contoso products aimed at individuals. Find Contoso products aimed at businesses. Find Contoso products aimed at individuals with the name mark8. Find Contoso products aimed at businesses with the name mark8."
        },
    }
    ```

    As we're going to be adding another parameter to the command, also update the command description to include the new parameter.

1. In the **commands** array, update the **description** of the command

    ```json
    {
        "commands": [
            {
                "id": "Search",
                "type": "query",
                "title": "Products",
                "description": "Find products by name or by target audience",
                "initialRun": true,
                "fetchTask": false,
                "context": [...],
                "parameters": [...]
            }
        ]
    }
    ```

    Now add a new parameter that Copilot can use. This new parameter helps users lookup up products using Copilot that is aimed at different audiences, such as individuals and businesses.

1. In the **parameters** array, add the **TargetAudience** parameter after the **ProductName** parameter.

    ```json
    {    
        "parameters": [
            {
                "name": "ProductName",
                "title": "Product name",
                "description": "The name of the product as a keyword",
                "inputType": "text"
            },
            {
                "name": "TargetAudience",
                "title": "Target audience",
                "description": "Audience that the product is aimed at. Consumer products are sold to individuals. Enterprise products are sold to businesses",
                "inputType": "text"
            }
        ]
    }
    ```

1. Save your changes

The description of the **TargetAudience** parameter describes what it's and explains that the parameter should accept **Consumer** or **Enterprise** are allowed values.

## Task 2 - Update message extension logic

To support the new parameter, and support complex prompts, update the OnTeamsMessagingExtensionQueryAsync method in the Bot Activity Handler to handle multiple parameters.

Say a user enters the prompt, "Find Contoso products aimed at individuals with the name Mark8". Given the parameter descriptions, "aimed at individuals" is translated to **Consumer** and passed as the value of the **TargetAudience** parameter. 'Mark8' is passed as the value of the **ProductName** parameter.

Continuing in Visual Studio:

1. In **Search** folder, open the file named **SearchApp.cs**
1. In the **OnTeamsMessagingExtensionQueryAsync** method, find the below code block:

    ```csharp
    var name = GetQueryData(query.Parameters, "ProductName");
    var nameFilter = !string.IsNullOrEmpty(name) ? $"startswith(fields/Title, '{name}')" : string.Empty;
    var filters = new List<string> { nameFilter };
    var filterQuery = filters.Count == 1 ? filters.FirstOrDefault() : string.Join(" and ", filters); 
    ```

1. Update the code block to get the value of the **TargetAudience** parameter and create a filter query to use when querying the SharePoint Online list.

    ```csharp
    var name = GetQueryData(query.Parameters, "ProductName");
    var retailCategory = GetQueryData(query.Parameters, "TargetAudience");
    
    var nameFilter = !string.IsNullOrEmpty(name) ? $"startswith(fields/Title, '{name}')" : string.Empty;
    var retailCategoryFilter = !string.IsNullOrEmpty(retailCategory) ? $"fields/RetailCategory eq '{retailCategory}'" : string.Empty;
    var filters = new List<string> { nameFilter };
    filters.RemoveAll(f => string.IsNullOrEmpty(f));
    var filterQuery = filters.Count == 1 ? filters.FirstOrDefault() : string.Join(" and ", filters);
    ```

1. Save your changes

## Task 3 - Provision resources

Run the Prepare Teams App Dependencies process to provision resources.

Continuing in Visual Studio:

1. In **Solution Explorer**, right-click the **MsgExtProductSupport** project
1. Expand the **Teams Toolkit** menu, select **Prepare Teams App Dependencies**
1. In the **Microsoft 365 account** dialog, select **Continue**
1. In the **Provision** dialog, select **Provision**
1. In the **Teams Toolkit warning** dialog, select **Provision**
1. In the **Teams Toolkit information** dialog, **close** the prompt

## Task 4 - Run and debug

Now start the web service and test the message extension in Copilot for Microsoft 365.

1. Press **F5** to start a debugging session and open a new browser window that navigates to the Microsoft Teams web client.
1. Enter your Microsoft 365 account credentials and continue to Microsoft Teams.
1. In the app install dialog, select **Add**
1. Open the **Copilot** app from Microsoft Teams
1. In the compose message area, open the **Plugins** flyout
1. In the list of plugins, toggle the **Contoso products** plugin to enable it
1. Enter **Find Contoso products aimed at individuals** as your message and send it
1. In the Copilot response, a sign in button is returned, select the **Sign in** button to authenticate
1. After successfully authenticating, **Enter Find Contoso products aimed at individuals** as your message and send it
1. In the Copilot response, data returned in the plugin response is displayed and the plugin referenced in the response
1. To view the Adaptive Card relevant to the result, hover over the references in the Copilot response

Close the browser to stop the debugging session.

[Continue to the lab summary...](./6-summary.md)