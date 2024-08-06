---
lab:
    title: 'Exercise 3 - Return product data from Microsoft Entra protected API'
    module: 'LAB 03: Connect Copilot for Microsoft 365 to your external data in real-time with message extension plugins built with .NET and Visual Studio'
---

# Exercise 3 - Return product data from Microsoft Entra protected API

In this exercise, you update the message extension to retrieve data from a custom API. You get data from the custom API based on the user query and return data in search results to the user.

![Screenshot of search results returned by a search based message extension in Microsoft Teams.](../media/3-search-results-api.png)

## Task 1 - Install and configure Dev Proxy

In this exercise, you use Dev Proxy, a command-line tool that can simulate APIs. It's useful when you want to test your app without having to create a real API.

To complete this exercise, you need to install the [latest version of of Dev Proxy](/microsoft-cloud/dev/dev-proxy/get-started) and download the Dev Proxy preset for this module.

The preset simulates a CRUD (Create, Read, Update, Delete) API with an in-memory data store, which is protected by Microsoft Entra. This means that you can test your app as if it were calling a real API that requires authentication.

1. To install Dev Proxy, run the following command in your terminal window:

    ```bash
    winget install Microsoft.DevProxy --silent
    ```

1. To download the preset, run the following command in your terminal (**Ctrl+`** in Visual Studio):

    ```bash
    devproxy preset get learn-copilot-me-plugin
    ```

## Task 2 - Get the user query value

Create a method that gets the user query value by the name of the parameter.

In Visual Studio and the **ProductsPlugin** project:

1. In the **Helpers** folder, create a new file named **MessageExtensionHelpers.cs**.

1. In the file, add the following code:

   ```csharp
   using Microsoft.Bot.Schema.Teams;

   internal class MessageExtensionHelpers
   {
       internal static string GetQueryParameterValueByName(IList<MessagingExtensionParameter> parameters, string name) => parameters.FirstOrDefault(p => p.Name == name)?.Value as string ?? string.Empty;
   }
   ```

1. Save your changes.

Next, update the **OnTeamsMessagingExtensionQueryAsync** method in the SearchApp class to use the new helper method.

1. In the **Search** folder, open **SearchApp.cs**.

1. In the **OnTeamsMessagingExtensionQueryAsync** method, replace the following code:

   ```csharp
   var text = query?.Parameters?[0]?.Value as string ?? string.Empty;
   ```

   with

   ```csharp
   var text = MessageExtensionHelpers.GetQueryParameterValueByName(query.Parameters, "ProductName");
   ```

1. Move the cursor to the **text** variable, use `Ctrl + R`, `Ctrl + R`, and rename the variable to **name**.

1. Press **Enter** to rename the variable across 3 files.

1. Save your changes.

The **OnTeamsMessagingExtensionQueryAsync** method should now look like this:

```csharp
protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
{
    var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
    var tokenResponse = await AuthHelpers.GetToken(userTokenClient, query.State, turnContext.Activity.From.Id, turnContext.Activity.ChannelId, connectionName, cancellationToken);

    if (!AuthHelpers.HasToken(tokenResponse))
    {
        return await AuthHelpers.CreateAuthResponse(userTokenClient, connectionName, (Activity)turnContext.Activity, cancellationToken);
    }

    var name = MessageExtensionHelpers.GetQueryParameterValueByName(query.Parameters, "ProductName");

    var card = await File.ReadAllTextAsync(Path.Combine(".", "Resources", "card.json"), cancellationToken);
    var template = new AdaptiveCardTemplate(card);

    return new MessagingExtensionResponse
    {
        ComposeExtension = new MessagingExtensionResult
        {
            Type = "result",
            AttachmentLayout = "list",
            Attachments = [
                new MessagingExtensionAttachment
                    {
                        ContentType = AdaptiveCard.ContentType,
                        Content = JsonConvert.DeserializeObject(template.Expand(new { title = name })),
                        Preview = new ThumbnailCard { Title = name }.ToAttachment()
                    }
            ]
        }
    };
}
```

## Task 3 - Get data from the custom API

To get data from the custom API, you need to send the access token in the Authorization header of the request and deserialize the response into a model that represents the product data.

First, create a model that represents the product data that is returned from the custom API.

In Visual Studio and the **ProductsPlugin** project:

1. Create a folder named **Models**.

1. In the **Models** folder, create a new file named **Product.cs**.

1. In the file, add the following code:

   ```csharp
   using System.Text.Json.Serialization;

   internal class Product
   {
       [JsonPropertyName("productId")]
       public int Id { get; set; }
       [JsonPropertyName("imageUrl")]
       public string ImageUrl { get; set; }
       [JsonPropertyName("name")]
       public string Name { get; set; }
       [JsonPropertyName("category")]
       public string Category { get; set; }
       [JsonPropertyName("callVolume")]
       public int CallVolume { get; set; }
       [JsonPropertyName("releaseDate")]
       public string ReleaseDate { get; set; }
   }
   ```

1. Save your changes.

Next, create a service class that retrieves the product data from the custom API.

1. Create a folder named **Services**.

1. In the **Services** folder, create a new file named **ProductService.cs**.

1. In the file, add the following code:

    ```csharp
    using System.Net.Http.Headers;
    
    internal class ProductsService
    {
        private readonly HttpClient _httpClient;
        private readonly string _baseUri = "https://api.contoso.com/v1/";
    
        internal ProductsService(string token)
        {
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }
    
        internal async Task<Product[]> GetProductsByNameAsync(string name)
        {
            var response = await _httpClient.GetAsync($"{_baseUri}products?name={name}");
            response.EnsureSuccessStatusCode();
            var jsonString = await response.Content.ReadAsStringAsync();
            return System.Text.Json.JsonSerializer.Deserialize<Product[]>(jsonString);
        }
    }
    ```

1. Save your changes.

The **ProductsService** class contains methods to get product data from the custom API. The class constructor takes an access token as a parameter and sets up an **HttpClient** instance with the access token in the Authorization header.

Next, update the **OnTeamsMessagingExtensionQueryAsync** method to use the **ProductsService** class to get product data from the custom API.

1. In the **Search** folder, open **SearchApp.cs**.

1. In the **OnTeamsMessagingExtensionQueryAsync** method, add the following code after **name** variable declaration to get product data from the custom API:

   ```csharp
   var productService = new ProductsService(tokenResponse.Token);
   var products = await productService.GetProductsByNameAsync(name);
   ```

1. Save your changes.

## Task 4 - Create search results

Now that you have the product data, you can include it in the search results that are returned to the user.

First, let's update the existing Adaptive Card template to display the product information.

Continuing in Visual Studio and in the **ProductsPlugin** project:

1. In the **Resources** folder, rename **card.json** to **Product.json**.

1. In the **Resources** folder, create a new file named **Product.data.json**. This file contains example data that Visual Studio uses to generate a preview of the Adaptive Card template.

1. In the file, add the following JSON:

    ```json
    {
      "callVolume": 36,
      "category": "Enterprise",
      "imageUrl": "https://raw.githubusercontent.com/SharePoint/sp-dev-provisioning-templates/master/tenant/productsupport/source/Product%20Imagery/Contoso4.png",
      "name": "Contoso Quad",
      "productId": 1,
      "releaseDate": "2019-02-09"
    }
    ```

1. Save your changes.

1. In the **Resources** folder, open **Product.json**.

1. In the file, replace the contents with the following JSON:

    ```json
    {
      "type": "AdaptiveCard",
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "version": "1.5",
      "body": [
        {
          "type": "TextBlock",
          "text": "${name}",
          "wrap": true,
          "style": "heading"
        },
        {
          "type": "TextBlock",
          "text": "${category}",
          "wrap": true
        },
        {
          "type": "Container",
          "items": [
            {
              "type": "Image",
              "url": "${imageUrl}",
              "altText": "${name}"
            }
          ],
          "minHeight": "350px",
          "verticalContentAlignment": "Center",
          "horizontalAlignment": "Center"
        },
        {
          "type": "FactSet",
          "facts": [
            {
              "title": "Call Volume",
              "value": "${formatNumber(callVolume,0)}"
            },
            {
              "title": "Release Date",
              "value": "${formatDateTime(releaseDate,'dd/MM/yyyy')}"
            }
          ]
        }
      ]
    }
    ```

1. Save your changes

The Adaptive Card template uses binding expressions to display the product information. The **\$\{name\}**, **\$\{category\}**, **\$\{imageUrl\}**, **\$\{callVolume\}**, and **\$\{releaseDate\}** expressions are replaced with the corresponding values from the product data. The **formatNumber** and **formatDateTime** template functions are used to format the **callVolume** and **releaseDate** values, into a number and date respectively.

Take a moment to explore the Adaptive Card preview in Visual Studio. The preview shows how the Adaptive Card template looks when the product data is bound to the template. It uses the example data from the **Product.data.json** file to generate the preview.

Next, update the **validDomains** property in the app manifest to include the **raw.githubusercontent.com** domain, so that the images in the Adaptive Card template can be displayed in Microsoft Teams.

In the **TeamsApp** project:

1. In the **appPackage** folder, open **manifest.json**.

1. In the file, add the GitHub domain to the **validDomains** property:

    ```json
      "validDomains": [
        "token.botframework.com",
        "raw.githubusercontent.com",
        "${{BOT_DOMAIN}}"
      ],
    ```

1. Save your changes.

Next, update the **OnTeamsMessagingExtensionQueryAsync** method to create a list of attachments that contain the product information.

In the **ProductsPlugin** project:

1. In the **Search** folder, open **SearchApp.cs**,

1. Update **card.json** to **Product.json**, to reflect the change in the file name. Replace the following code:

   ```csharp
   var card = await File.ReadAllTextAsync(Path.Combine(".", "Resources", "card.json"), cancellationToken);
   ```

   with

   ```csharp
   var card = await File.ReadAllTextAsync(Path.Combine(".", "Resources", "Product.json"), cancellationToken);
   ```

1. Add the following code after the **template** variable declaration to create a list of attachments:

   ```csharp
    var attachments = products.Select(product =>
    {
        var content = template.Expand(product);
    
        return new MessagingExtensionAttachment
        {
            ContentType = AdaptiveCard.ContentType,
            Content = JsonConvert.DeserializeObject(content),
            Preview = new ThumbnailCard
            {
                Title = product.Name,
                Subtitle = product.Category,
                Images = [new() { Url = product.ImageUrl }]
            }.ToAttachment()
        };
    }).ToList();
   ```

1. Update the **return** statement to include the **attachments** variable:

   ```csharp
   return new MessagingExtensionResponse
   {
       ComposeExtension = new MessagingExtensionResult
       {
           Type = "result",
           AttachmentLayout = "list",
           Attachments = attachments
       }
   };
   ```

1. Save changes

## Task 5 - Create and update resources

With everything now in place, run the **Prepare Teams App Dependencies** process to create new resources and update existing ones.

Continuing in Visual Studio:

1. In **Solution Explorer**, right-click the **TeamsApp** project,

1. Expand the **Teams Toolkit** menu, select **Prepare Teams App Dependencies**,

1. In the **Microsoft 365 account** dialog, select **Continue**,

1. In the **Provision** dialog, select **Provision**,

1. In the **Teams Toolkit warning** dialog, select **Provision**,

1. In the **Teams Toolkit information** dialog, select the cross icon to close the dialog,

## Task 6 - Run and debug

With the resources provisioned, start a debugging session to test the message extension.

First, start Dev Proxy to simulate the custom API.

1. Open a new PowerShell or Terminal window as Administrator.

1. Run the following command to start Dev Proxy:

   ```bash
   devproxy --config-file "~appFolder/presets/learn-copilot-me-plugin/products-api-config.json"
   ```

1. If prompted, accept the certificate warning

> [!NOTE]
> When Dev Proxy is running, it acts as a system-wide proxy.

Next, start a debug session in Visual Studio:

1. To start a new debug session, press <kbd>F5</kbd> or select **Start** from the toolbar.

1. Wait until a browser window opens and the app install dialog appears in the Microsoft Teams web client. If prompted, enter your Microsoft 365 account credentials.

1. In the app install dialog, select **Add**.

1. Open a new, or existing Microsoft Teams chat.

1. In the message compose area, type **/apps** to open the app picker.

1. In the list of apps, select **Contoso products** to open the message extension.

1. In the text box, enter **mark8**.

1. Wait for the search to complete and the results to be displayed.

    ![Screenshot of search results returned by a search based message extension in Microsoft Teams.](../media/3-search-results-api.png)

1. In the list of results, select a search result to embed a card into the compose message box

Return to Visual Studio and select **Stop** from the toolbar, or press <kbd>Shift</kbd> + <kbd>F5</kbd> to stop the debug session. Also, shut down Dev Proxy using <kbd>Ctrl</kbd> + <kbd>C</kbd>.

[Continue to the next exercise...](./5-exercise-extend-optimize-message-extensions-copilot-microsoft-365.md)