---
lab:
    title: 'Exercise 3 - Retrieve product information from SharePoint Online'
    module: 'LAB 03: Connect Copilot for Microsoft 365 to your external data in real-time with message extension plugins built with .NET and Visual Studio'
---

# Exercise 3 - Retrieve product information from SharePoint Online

In this exercise, you provision and configure a SharePoint Online site, which stores product information as items in a list. You update the message extension code to retrieve the list items from SharePoint Online using the Microsoft Graph SDK and return the list item data in the search results. Finally, you run and debug your message extension and test it in Microsoft Teams.

![Screenshot of search results returned by a search based message extension in Microsoft Teams. The search results are returned from SharePoint Online. Each search result displays the product name, category, and product image.](../media/4-search-results-sharepoint-online.png)

## Task 1 - Provision and configure the Product marketing SharePoint site

Start by creating a SharePoint Online site using the SharePoint look book service.

In a web browser:

1. Go to **SharePoint look book** at [https://lookbook.microsoft.com](https://lookbook.microsoft.com)
1. On the top navigation, expand **View the designs**
1. In the **View the designs** menu, expand **Team** and select **Product Support**
1. Select **Add you your tenant**
1. If prompted, sign in to your tenant.
1. On the permission consent screen, review the required permissions and select **Accept** to return to the SharePoint look book service.
1. In the form, accept the defaults and select **Provision**

An email is sent to your email address to notify you when the site provisioning is complete. This process can take a few minutes to complete.

![Screenshot of the homepage from the Product support SharePoint Online team site. A list of recently released products is shown.](../media/1-sharepoint-online-product-support-site.png)

To enable filtering on the Title and Retail Category columns when querying the list using Microsoft Graph API, create indexes on the list.

Continuing in the web browser:

1. Go to the **Product support** site at **<https://tenant.sharepoint.com/sites/productmarketing>**, replace **tenant** with the name of your SharePoint Online instance
1. On the **Microsoft 365 suite bar**, select the **settings cog** to open the Settings side panel.
1. Under the **SharePoint** heading, select **Site contents**
1. In the list of Lists and Libraries, hover over the **Products** list, select the **vertical three dots** icon to expand the **Show actions** menu, then select **Settings**.
1. In the **Columns** section, under the list of columns, select **Indexed columns**
1. Select **Create a new index**

## Task 2 - Add SharePoint host name and Site URL environment variables

Next, we centralize the hostname of your SharePoint Online instance and the Product support site URL as environment variables. You then expose the values as an environment variable to use at run time and update the code to read the value.

Open Visual Studio:

1. In the **env** folder, open the file named **.env.local**
1. In the file, add the **SPO_HOSTNAME** and **SPO_SITE_URL** environment variables, replacing **tenant** with the name of your SharePoint Online instance:

    ```text
    SPO_HOSTNAME=tenant.sharepoint.com
    SPO_SITE_URL=sites/productmarketing
    ```

1. Save your changes

Next, update the action to write the environment variables to the app settings file.

1. In the project root folder, open the file named **teamsapp.local.yml**
1. Find the step that uses the **file/createOrUpdateJsonFile** action, which targets the **./appsettings.Development.json** file
1. In the file, update the **content** array, adding the **SPO_HOSTNAME** and **SPO_SITE_URL** variables:

    ```yml
      - uses: file/createOrUpdateJsonFile
        with:
          target: ./appsettings.Development.json
          content:
            BOT_ID: ${{BOT_ID}}
            BOT_PASSWORD: ${{SECRET_BOT_PASSWORD}}
            CONNECTION_NAME: ${{CONNECTION_NAME}}
            SPO_HOSTNAME: ${{SPO_HOSTNAME}}
            SPO_SITE_URL: ${{SPO_SITE_URL}}
    ```

1. Save your changes

Now, update the ConfigOptions class to include the new environment variables

1. In the project root folder, open Config.cs
1. In the ConfigOptions class, add new string properties with the names SPO_HOSTNAME and SPO_SITE_URL

    ```csharp
    public class ConfigOptions
    {
      public string BOT_ID { get; set; }
      public string BOT_PASSWORD { get; set; }
      public string CONNECTION_NAME { get; set; }
      public string SPO_HOSTNAME { get; set; }
      public string SPO_SITE_URL { get; set; }
    }
    ```

1. Save your changes

Next, update the app configuration with the two environment variables.

1. In the project root folder, open Program.cs
1. Add new lines to add the SPO_HOSTNAME and SPO_SITE_URL environment variables as app configuration settings.

    ```csharp
    var config = builder.Configuration.Get<ConfigOptions>();
    builder.Configuration["MicrosoftAppType"] = "MultiTenant";
    builder.Configuration["MicrosoftAppId"] = config.BOT_ID;
    builder.Configuration["MicrosoftAppPassword"] = config.BOT_PASSWORD;
    builder.Configuration["CONNECTION_NAME"] = config.CONNECTION_NAME;
    builder.Configuration["SPO_HOSTNAME"] = config.SPO_HOSTNAME;
    builder.Configuration["SPO_SITE_URL"] = config.SPO_SITE_URL;
    ```

1. Save your changes

The last step is to update the Bot Activity Handler to read the values from the app configuration and store the value in read-only properties.

1. In the Search folder, open the file named SearchApp.cs
1. In the SearchApp class, create read-only string properties with the names spoHostname and spoSiteUrl

    ```csharp
    public class SearchApp : TeamsActivityHandler
    {
      private readonly string connectionName;
      private readonly string spoHostname;
      private readonly string spoSiteUrl;
    }
    ```

1. Update the constructor to set the property values using the injected app configuration:

    ```csharp
    public SearchApp(IConfiguration configuration)
    {
        connectionName = configuration["CONNECTION_NAME"];
        spoHostname = configuration["SPO_HOSTNAME"];
        spoSiteUrl = configuration["SPO_SITE_URL"];
    } 
    ```

1. Save your changes.

## Task 3 - Update search command

As the message extension returns product information, update the search command title and description, also update the parameter name and its description.

Continuing in Visual Studio:

1. In the **appPackage** folder, open the file named **manifest.json**
1. In the **composeExtensions** array, update the command object with:

    ```json
    "composeExtensions": [
      {
        "botId": "${{BOT_ID}}",
        "commands": [
          {
            "id": "Search",
            "type": "query",
            "title": "Products",
            "description": "Find products by name",
            "initialRun": false,
            "fetchTask": false,
            "context": [
              "commandBox",
              "compose",
              "message"
            ],
            "parameters": [
              {
                "name": "ProductName",
                "title": "Product name",
                "description": "The name of the product as a keyword",
                "inputType": "text"
              }
            ]
          }
        ]
      }
    ],
    ```

1. Save your changes

## Task 4 - Get the user query value

When the OnTeamsMessagingExtensionQueryAsync method is executed, the first thing we want to do is understand what the user entered into the search box.

First, let’s remove the existing code.

Continuing in Visual Studio:

1. In the **Search** folder, open the file named **SearchApp.cs**
1. In the **OnTeamsMessagingExtensionQueryAsync** method, remove all the code **after** the **if** statement that checks for an access token
1. In the **SearchApp** class, remove the **FindPackages** method and the **_adaptiveCardFilePath** property.

After you remove the existing code, your **SearchApp** class should look like the following code snippet:

```csharp
public class SearchApp : TeamsActivityHandler
{
    private readonly string connectionName;
    private readonly string spoHostname;
    private readonly string spoSiteUrl;

    public SearchApp(IConfiguration configuration)
    {
        connectionName = configuration["CONNECTION_NAME"];
        spoHostname = configuration["SPO_HOSTNAME"];
        spoSiteUrl = configuration["SPO_SITE_URL"];
    }

    protected override async Task<MessagingExtensionResponse> OnTeamsMessagingExtensionQueryAsync(ITurnContext<IInvokeActivity> turnContext, MessagingExtensionQuery query, CancellationToken cancellationToken)
    {
        var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
        var tokenResponse = await GetToken(userTokenClient, query.State, turnContext.Activity.From.Id, turnContext.Activity.ChannelId, connectionName, cancellationToken);

        if (!HasToken(tokenResponse))
        {
            return await CreateAuthResponse(userTokenClient, connectionName, (Activity)turnContext.Activity, cancellationToken);
        }
    }
}
```

Now let’s write the code to get the value of the **ProductName** parameter.

1. In the **OnTeamsMessagingExtensionQueryAsync** method, add code to retrieve the value of the **ProductName** parameter from the **Parameters** array in the **MessagingExtensionQuery** object

    ```csharp
    var name = GetQueryData(query.Parameters, "ProductName");
    ```

1. In the **SearchApp** class, implement the **GetQueryData** method

    ```csharp
    private static string GetQueryData(IList<MessagingExtensionParameter> parameters, string key)
    {
      if (parameters.Any() != true)
      {
        return string.Empty;
      }
    
      var foundPair = parameters.FirstOrDefault(pair => pair.Name == key);
      return foundPair?.Value?.ToString() ?? string.Empty;
    }
    ```

1. Save your changes

The **GetQueryData** method is used to retrieve the value associated with a specific key from a list of **MessagingExtensionParameter** objects. It provides a convenient way to extract data from the parameters array in the **MessagingExtensionQuery** object.

## Task 5 - Create SharePoint list query OData filter

Now that we have the value passed in by the user, use this value to create an OData query filter. The filter is used to query the SharePoint Online list by the Title column, which contains the product name.

Continuing in Visual Studio:

1. In the **OnTeamsMessagingExtensionQueryAsync** method, add code to create the **filterQuery** variable.

    ```csharp
    var nameFilter = !string.IsNullOrEmpty(name) ? $"startswith(fields/Title, '{name}')" : string.Empty;
    var filters = new List<string> { nameFilter };
    var filterQuery = filters.Count == 1 ? filters.FirstOrDefault() : string.Join(" and ", filters);
    ```

1. Save your changes

The code constructs a filter query based on the **name** parameter. If the name parameter is provided, it creates a filter expression to search for items with a **Title** field starting with the provided name. If the name parameter isn't provided, it assigns an empty string as the filter query. The resulting filter query is used later in the code to retrieve filtered items from a SharePoint site.

## Task 6 - Install and configure Microsoft Graph SDK

To execute authenticated requests to Microsoft Graph, you use the **Microsoft Graph SDK**.

Install the Microsoft Graph SDK package from NuGet, then create a **TokenProvider** class, which allows you to use the access token obtained from the token service and then initialize a new **GraphServiceClient**.

Continuing in Visual Studio:

1. In Solution Explorer, right-click **MsgExtProductSupport** project
1. Select **Manage NuGet Packages…**
1. Select the **Browse** tab and search for **Microsoft.Graph**
1. In the list of results, select **Microsoft.Graph**
1. In the **Version** dropdown, select **5.42.0**
1. Select **Install**
1. In the **License Acceptance** dialog box, select **I Accept** to install the SDK

After installing the package, create a token provider for the Microsoft Graph SDK.

1. In the **Search** folder, create a new file named **TokenProvider.cs**
1. In the file, add the following code:

    ```csharp
    using Microsoft.Kiota.Abstractions.Authentication;
    
    namespace MsgExtProductSupport.Search
    {
       public class TokenProvider : IAccessTokenProvider
        {
            public string Token { get; set; }
            public AllowedHostsValidator AllowedHostsValidator => throw new NotImplementedException();
    
            public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object>? additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
            {
                return Task.FromResult(Token);
            }
        }
    }
    ```

1. Save your changes

Now create a method to create a new **GraphServiceClient** instance.

1. In the **Search** folder, open the file named **SearchApp.cs**
1. In the file, import the namespaces that you need:

    ```csharp
    using Microsoft.Graph;
    using Microsoft.Kiota.Abstractions.Authentication;
    ```

1. In the **OnTeamsMessagingExtensionQueryAsync** method, add code to create a new Graph Client, which you use to send requests to Microsoft Graph

    ```csharp
    var graphClient = CreateGraphClient(tokenResponse);
    ```

1. In the **SearchApp** class, implement the **CreateGraphClient** method

    ```csharp
    private static GraphServiceClient CreateGraphClient(TokenResponse tokenResponse)
    {
      TokenProvider provider = new() { Token = tokenResponse.Token };
      var authenticationProvider = new BaseBearerTokenAuthenticationProvider(provider);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }
    ```

1. Save your changes

This code sets up the authentication provider and creates a client object that can be used to interact with the Microsoft Graph API using the provided access token.

## Task 7 - Query products list

To query the items in the Products list and later, create the search results, you use the GraphServiceClient to send requests to get product data from SharePoint Online.

Continuing in Visual Studio:

In the **OnTeamsMessagingExtensionQueryAsync** method, add code to get the SharePoint data:

  ```csharp
  var site = await GetSharePointSite(graphClient, spoHostname, spoSiteUrl, cancellationToken);
  var drive = await GetSharePointDrive(graphClient, site.SharepointIds.SiteId, "Product Imagery", cancellationToken);
  var items = await GetProducts(graphClient, site.SharepointIds.SiteId, filterQuery, cancellationToken);
  ```

This code will:

- **Get the Product Marketing site**, the site object contains the ID of the SharePoint site, which is used to return and query objects in the site
- **Get the Product Imagery drive**, the drive represents the document library that contains the product images. You use the drive later to get the product images that you display in the search results
- **Get the Products**, which you use to query the Products list based on the user query

Implement the three methods in the **SearchApp** class.

- Implement the **GetSharePointSite** method

    ```csharp
    private static async Task<Site> GetSharePointSite(GraphServiceClient graphClient, string hostName, string siteUrl, CancellationToken cancellationToken)
    {
        return await graphClient.Sites[$"{hostName}:/{siteUrl}"].GetAsync(r => r.QueryParameters.Select = new string[] { "sharePointIds" }, cancellationToken);
    }
    ```

This method uses the GraphServiceClient to send a request to Microsoft Graph to return the site object using a path. The path is created from combining the SharePoint Online hostname and site URL. As you only need the sharePointIds property values, the Select query parameter is configured to only return this property in the response.

- Implement the **GetSharePointDrive** method

    ```csharp
    private static async Task<Drive> GetSharePointDrive(GraphServiceClient graphClient, string siteId, string name, CancellationToken cancellationToken)
    {
        var drives = await graphClient.Sites[siteId].Drives.GetAsync(r => r.QueryParameters.Select = new string[] { "id", "name" }, cancellationToken);
        var drive = drives.Value.Find(d => d.Name == name);
        return drive;
    }
    ```

This method uses the GraphServiceClient and the site ID to return a collection of document libraries from the site, returning the ID and name properties of each document library. The collection of libraries is then filtered to return drive, which has the same name as the name method parameter.

- Implement the **GetProducts** method

    ```csharp
    private static async Task<SiteCollectionResponse> GetProducts(GraphServiceClient graphClient, string siteId, string filterQuery, CancellationToken cancellationToken)
    {
        var fields = new string[]
        {
            "fields/Id",
            "fields/Title",
            "fields/RetailCategory",
            "fields/PhotoSubmission",
            "fields/CustomerRating",
            "fields/ReleaseDate"
        };
    
        var request = graphClient.Sites.WithUrl($"https://graph.microsoft.com/v1.0/sites/{siteId}/lists/Products/items?expand={string.Join(",", fields)}&$filter={filterQuery}");
        return await request.GetAsync(null, cancellationToken);
    }
    ```

This method uses the GraphServiceClient to return filtered list items from the Products list using the passed in filter query and returns the list item data defined in the fields array.

## Task 8 - Create search results

After getting the products from SharePoint, you’ll create the search results, which will be returned to the user.

Creating the search results consists of iterating over the items array, for each item, you create a MessagingExtensionAttachment, which contains the preview and content cards.

Before you iterate over the items, create an Adaptive Card template, which you use in your loop.

Continuing in Visual Studio:

1. In the **Resources** folder, create a new file named **Product.json**

    ```json
    {
      "type": "AdaptiveCard",
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "version": "1.6",
      "body": [
        {
          "type": "TextBlock",
          "text": "${Product.Title}",
          "wrap": true,
          "style": "heading"
        },
        {
          "type": "TextBlock",
          "text": "${Product.RetailCategory}",
          "wrap": true
        },
        {
          "type": "Container",
          "items": [
            {
              "type": "Image",
              "url": "${ProductImage}",
              "altText": "${Product.Title}"
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
              "value": "${formatNumber(Product.CustomerRating,0)}"
            },
            {
              "title": "Release Date",
              "value": "${formatDateTime(Product.ReleaseDate,'dd/MM/yyyy')}"
            }
          ]
        },
        {
          "type": "ActionSet",
          "actions": [
            {
              "type": "Action.OpenUrl",
              "title": "View",
              "url": "https://${SPOHostname}/${SPOSiteUrl}/Lists/Products/DispForm.aspx?ID=${Product.Id}"
            }
          ]
        }
      ]
    }
    ```

The template uses data binding expressions, which are replaced with actual values when rendering the Adaptive Card. When the card is rendered, it contains some product information, a product image, and an action button. The action button opens a browser that navigates to the SharePoint List display form for the product item.

Next, add the code to transform the JSON file into an Adaptive Card template.

1. In the **Search** folder, open **SearchApp.cs**
1. In the file, import the namespaces that you need:

    ```csharp
    using AdaptiveCards.Templating;
    ```

1. In the **OnTeamsMessagingExtensionQueryAsync** method, add the code to read the contents of the JSON file and create a new **AdaptiveCardTemplate** object.

    ```csharp
    var card = File.ReadAllText(@"Resources\Product.json");
    var template = new AdaptiveCardTemplate(card);
    ```

1. Save your changes

Next, you create a loop to iterate over the list items. Each iteration will:

- Deserialize the current item data into a Product object
- Get the Product image thumbnail
- Create a content card
- Create a preview card
- Create a MessagingExtensionAttachment combining the content and preview cards
- Add the MessagingExtensionAttachment to the list

Once the loop is finished; you have a list of attachments that can be returned to the user.

1. In the **Search** folder, open **SearchApp.cs**
1. In the **OnTeamsMessagingExtensionQueryAsync** method, add code to create a new List, to store **MessagingExtensionAttachment** objects in

    ```csharp
    var attachments = new List<MessagingExtensionAttachment>();
    ```

1. Create a foreach loop to iterate over the list items.

    ```csharp
    foreach (var item in items.Value) { 
            
    }
    ```

1. Add the following code to the foreach loop.

    ```csharp
    var product = JsonConvert.DeserializeObject<Product>(item.AdditionalData["fields"].ToString());
    product.Id = item.Id;
    
    var thumbnails = await GetThumbnails(graphClient, drive.Id, product.PhotoSubmission, cancellationToken);
    
    var resultCard = template.Expand(new
    {
      Product = product,
      ProductImage = thumbnails.Large.Url,
      SPOHostname = spoHostname,
      SPOSiteUrl = spoSiteUrl,
    });
    
    var previewcard = new ThumbnailCard
    {
      Title = product.Title,
      Subtitle = product.RetailCategory,
      Images = new List<CardImage> { new() { Url = thumbnails.Small.Url } }
    }.ToAttachment();
    
    var attachment = new MessagingExtensionAttachment
    {
      Content = JsonConvert.DeserializeObject(resultCard),
      ContentType = AdaptiveCard.ContentType,
      Preview = previewcard
    };
    
    attachments.Add(attachment);
    ```

1. Save your changes

To strongly type the list item data, create a model that represents the **Product**.

1. In the project root folder, create a new folder named **Models**
1. In the **Models** folder, create a new file called **Product.cs**

    ```csharp
    namespace MsgExtProductSupport.Models
    {
        public class Product
        {
            public string Title { get; set; }
            public string RetailCategory { get; set; }
            public Link Specguide { get; set; }
            public string PhotoSubmission { get; set; }
            public double CustomerRating { get; set; }
            public DateTime ReleaseDate { get; set; }
            public string Id { get; set; }
            public string ContentType { get; set; }
            public DateTime Modified { get; set; }
            public DateTime Created { get; set; }
        }
    
        public class Link
        {
            public string Description { get; set; }
            public string Url { get; set; }
        }
    }
    ```

1. Save your changes

Next, implement the **GetThumbnails** method to retrieve thumbnails images from Microsoft Graph for a product.

1. In the **Search** folder, open the file named **SearchApp.cs**
1. In the **SearchApp** class, create the **GetThumbnails** method

    ```csharp
    private static async Task<ThumbnailSet> GetThumbnails(GraphServiceClient graphClient, string driveId, string photoUrl, CancellationToken cancellationToken)
    {
        var fileName = photoUrl.Split('/').Last();
        var driveItem = await graphClient.Drives[driveId].Root.ItemWithPath(fileName).GetAsync(null, cancellationToken);
        var thumbnails = await graphClient.Drives[driveId].Items[driveItem.Id].Thumbnails["0"].GetAsync(r => r.QueryParameters.Select = new string[] { "small", "large" }, cancellationToken);
        return thumbnails;
    }
    ```

1. Save your changes

The **GetThumbnails** method uses the Thumbnails endpoint in the Microsoft Graph API to return small and large thumbnails of the product image stored in SharePoint.

## Task 9 - Return search results

Now that we have a collection of MessagingExtensionResult objects we can return them back to the user as search results.

- In the **OnTeamsMessagingExtensionQueryAsync** method, add code to return the search results as the message extension response.

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

## Task 10 - Provision resources

Run the Prepare Teams App Dependencies process to provision resources.

Continuing in Visual Studio:

1. In **Solution Explorer**, right-click the **MsgExtProductSupport** project
1. Expand the **Teams Toolkit** menu, select **Prepare Teams App Dependencies**
1. In the **Microsoft 365 account** dialog, select **Continue**
1. In the **Provision** dialog, select **Provision**
1. In the **Teams Toolkit warning** dialog, select **Provision**
1. In the **Teams Toolkit information** dialog, **close** the prompt

## Task 11 - Run and debug

Now start the web service and test the message extension in Microsoft Teams.

Continuing in Visual Studio:

1. Press **F5** to start a debugging session and open a new browser window that navigates to the Microsoft Teams web client.
1. Enter your Microsoft 365 account credentials and continue to Microsoft Teams.
1. In the app install dialog, select **Add**
1. Open a new, or existing Microsoft Teams chat
1. In the message compose area, select **…** to open the app flyout
1. In the list of apps, select **Contoso products** to open the message extension
1. Enter **Mark8** into the text box. Two results are shown, **Mark8** and **Mark8 controller**
1. Select **Mark8** to embed a card into the compose message box
1. **Send the message** containing the card
1. In the submitted card, select the **View** button to view the SharePoint list item for the product in the Products list in a new tab

Close the browser to stop the debugging session.

[Continue to the next exercise...](./5-exercise-extend-optimize-message-extensions-copilot-microsoft-365.md)