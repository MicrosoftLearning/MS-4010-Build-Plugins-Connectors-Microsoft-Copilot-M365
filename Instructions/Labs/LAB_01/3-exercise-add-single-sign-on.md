---
lab:
    title: 'Exercise 2 - Add single sign-on'
    module: 'LAB 03: Connect Copilot for Microsoft 365 to your external data in real-time with message extension plugins built with .NET and Visual Studio'
---

# Exercise 2 - Add single sign-on

In this exercise, you update the message extension so that users are prompted to sign-in and authenticate. You configure the Bot Microsoft Entra app registration and app manifest to enable single sign-on. You configure a Microsoft Entra app registration to authenticate with Microsoft Graph and update the message extension logic to obtain an access token using the Bot Framework Token Service. Then you run and debug your message extension to test in Microsoft Teams.

## Task 1 - Configure single sign-on

First, you configure the Bot Microsoft Entra app registration.

In Visual Studio:

1. In the **infra\entra** folder, open the file named **entra.bot.manifest.json**
1. In the file, update the **identifierUris** array to set the Application ID URI

    ```json
    "identifierUris": [
        "api://${{BOT_DOMAIN}}/botid-${{BOT_ID}}"
    ]
    ```

1. In the file, update the **oauth2Permissions** array to create a scope to allow Teams to call web APIs as an admin or user:

    ```json
      "oauth2Permissions": [
        {
          "adminConsentDescription": "Allows Teams to call the app's web APIs as the current user.",
          "adminConsentDisplayName": "Teams can access app's web APIs",
          "id": "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}",
          "isEnabled": true,
          "type": "User",
          "userConsentDescription": "Enable Teams to call this app's web APIs with the same rights that you have",
          "userConsentDisplayName": "Teams can access app's web APIs and make requests on your behalf",
          "value": "access_as_user"
        }
      ]
    ```

1. In the file, update the **preAuthorizedApplications** array to add Microsoft Teams, Microsoft Outlook, and Copilot for Microsoft 365 clients to the list of authorized clients:

    ```json
      "preAuthorizedApplications": [
        {
          "appId": "1fec8e78-bce4-4aaf-ab1b-5451cc387264",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        },
        {
          "appId": "5e3ce6c0-2b1f-4285-8d4b-75ee78787346",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        },
        {
          "appId": "4765445b-32c6-49b0-83e6-1d93765276ca",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        },
        {
          "appId": "0ec893e0-5785-4de6-99da-4ed124e5296c",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        },
        {
          "appId": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        },
        {
          "appId": "bc59ab01-8403-45c6-8796-ac3ef710b3e3",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        },
        {
          "appId": "27922004-5251-4030-b22d-91ecd9a37ea4",
          "permissionIds": [
            "${{AAD_APP_ACCESS_AS_USER_PERMISSION_ID}}"
          ]
        }
      ]
    ```

1. Save your changes

Next, update the app manifest file to define the resource that the client should use when initiating a single sign-on flow in the app.

Continuing in Visual Studio:

1. In the **appPackage** folder, open the file named **manifest.json**
1. In the file, add the following code:

    ```json
    "webApplicationInfo": {
      "id": "${{BOT_ID}}",
      "resource": "api://${{BOT_DOMAIN}}/botid-${{BOT_ID}}"
    }
    ```

1. Save your changes

## Task 2 - Create Microsoft Entra app registration manifest file for Microsoft Graph

To authenticate with Microsoft Graph, create a new app registration manifest file.

Continuing in Visual Studio:

1. In the **infra\entra** folder, create a file named **entra.graph.manifest.json**
2. In the file, add the following code:

    ```json
    {
      "id": "${{GRAPH_ENTRA_APP_OBJECT_ID}}",
      "appId": "${{GRAPH_ENTRA_APP_ID}}",
      "name": "${{APP_INTERNAL_NAME}}-graph-${{TEAMSFX_ENV}}",
      "accessTokenAcceptedVersion": 2,
      "signInAudience": "AzureADMultipleOrgs",
      "optionalClaims": {
        "idToken": [],
        "accessToken": [
          {
            "name": "idtyp",
            "source": null,
            "essential": false,
            "additionalProperties": []
          }
        ],
        "saml2Token": []
      },
      "requiredResourceAccess": [
        {
          "resourceAppId": "Microsoft Graph",
          "resourceAccess": [
            {
              "id": "Sites.ReadWrite.All",
              "type": "Scope"
            }
          ]
        }
      ],
      "oauth2Permissions": [],
      "preAuthorizedApplications": [],
      "identifierUris": [],
      "replyUrlsWithType": [
        {
          "url": "https://token.botframework.com/.auth/web/redirect",
          "type": "Web"
        }
      ]
    }
    ```

1. Save your changes

The **requiredResourceAccess** array defines the API permission scopes and the **replyUrlsWithType** array defines the redirect URIs on the app registration.

Now update the project file with actions to create the app registration.

1. In the project root folder, open **teamsapp.local.yml**
1. In the file, find the step that uses the **arm/deploy** action
1. Before the step, add the following code:

    ```yml
      - uses: aadApp/create
        with:
          name: ${{APP_INTERNAL_NAME}}-graph-${{TEAMSFX_ENV}}
          generateClientSecret: true
          signInAudience: AzureADMultipleOrgs
        writeToEnvironmentFile:
          clientId: GRAPH_ENTRA_APP_ID
          clientSecret: SECRET_GRAPH_ENTRA_APP_CLIENT_SECRET
          objectId: GRAPH_ENTRA_APP_OBJECT_ID
          tenantId: GRAPH_ENTRA_APP_TENANT_ID
          authority: GRAPH_ENTRA_APP_OAUTH_AUTHORITY
          authorityHost: GRAPH_ENTRA_APP_OAUTH_AUTHORITY_HOST
    
      - uses: aadApp/update
        with:
          manifestPath: "./infra/entra/entra.graph.manifest.json"
          outputFilePath : "./build/entra.graph.manifest.${{TEAMSFX_ENV}}.json"
    ```

1. Save your changes

## Task 3 - Create OAuth connection setting

Azure AI Bot Service connection settings are used for managing user authentication in bots and message extensions.

First, centralize the name of the OAuth connection setting that is used to create the connection setting as an environment variable, then add code to use the environment variable value at run time.

Continuing in Visual Studio:  

1. In the **env** folder, open **.env.local**
1. In the file, add the following code:

    ```text
    CONNECTION_NAME=MicrosoftGraph
    ```

1. In the project root folder, open the file named **teamsapp.local.yml**
1. In the file, find the step that uses the **file/createOrUpdateJsonFile** action targeting the **./appsettings.Development.json** file
1. Update the content array, adding the **CONNECTION_NAME** variable

    ```yml
      - uses: file/createOrUpdateJsonFile
        with:
          target: ./appsettings.Development.json
          content:
            BOT_ID: ${{BOT_ID}}
            BOT_PASSWORD: ${{SECRET_BOT_PASSWORD}}
            CONNECTION_NAME: ${{CONNECTION_NAME}}
    ```

1. Save your changes
1. In the project root folder, open **Config.cs**
1. In the **ConfigOptions** class, add a new string property with the name **CONNECTION_NAME**

    ```csharp
    public class ConfigOptions
    {
      public string BOT_ID { get; set; }
      public string BOT_PASSWORD { get; set; }
      public string CONNECTION_NAME { get; set; }
    }
    ```

1. Save your changes
1. In the project root folder, open the file named **Program.cs**
1. In the file, add a new line to add the **CONNECTION_NAME** environment variable as an app configuration setting:

    ```csharp
    var config = builder.Configuration.Get<ConfigOptions>();
    builder.Configuration["MicrosoftAppType"] = "MultiTenant";
    builder.Configuration["MicrosoftAppId"] = config.BOT_ID;
    builder.Configuration["MicrosoftAppPassword"] = config.BOT_PASSWORD;
    builder.Configuration["CONNECTION_NAME"] = config.CONNECTION_NAME;
    ```

Next, update the Bot Activity Handler to access the app configuration.

1. In the **Search** folder, open the file named **SearchApp.cs**
1. In the **SearchApp** class, create a read-only string property with the name **connectionName**.

    ```csharp
    public class SearchApp : TeamsActivityHandler
    {
      private readonly string connectionName;
    }
    ```

1. Create a constructor, which injects the app configuration and sets the value of the connectionName property.

    ```csharp
    public class SearchApp : TeamsActivityHandler
    {
      private readonly string connectionName;
    
      public SearchApp(IConfiguration configuration)
      {
        connectionName = configuration["CONNECTION_NAME"];
      }  
    }
    ```

1. Save your changes

Next, you update the Bicep files to provision the OAuth connection setting.

First, update the parameters file to pass the credentials of the Microsoft Graph Microsoft Entra app registration and the name of the connection setting.

1. In the infra **folder**, open the file named **azure.parameters.local.json**
1. In the parameters **object**, add the **graphEntraAppClientId**, **graphEntraAppClientSecret**, and **connectionName** parameters

    ```json
    {
      "$schema": "https://schema.management.azure.com/schemas/2015-01-01/deploymentParameters.json#",
      "contentVersion": "1.0.0.0",
      "parameters": {
        "resourceBaseName": {
          "value": "bot-${{RESOURCE_SUFFIX}}-${{TEAMSFX_ENV}}"
        },
        "botEntraAppClientId": {
          "value": "${{BOT_ID}}"
        },
        "botDisplayName": {
          "value": "${{APP_DISPLAY_NAME}}"
        },
        "botAppDomain": {
          "value": "${{BOT_DOMAIN}}"
        },
        "graphEntraAppClientId": {
          "value": "${{GRAPH_ENTRA_APP_ID}}"
        },
        "graphEntraAppClientSecret": {
          "value": "${{SECRET_GRAPH_ENTRA_APP_CLIENT_SECRET}}"
        },
        "connectionName": {
          "value": "${{CONNECTION_NAME}}"
        }
      }
    }
    ```

1. Save your changes

Next, update the Bicep file.

1. In the **infra** folder, open the file named **azure.local.bicep**
1. In the file, after the **botAppDomain** parameter declaration, add the **graphEntraAppClientId**, **graphEntraAppClientSecret**, and **connectionName** parameter declarations

    ```bicep
    param graphEntraAppClientId string
    @secure()
    param graphEntraAppClientSecret string
    param connectionName string
    ```

1. In the **azureBotRegistration** module, add the **graphEntraAppClientId**, **graphEntraAppClientSecret**, and **connectionName** parameters

    ```bicep
    module azureBotRegistration './botRegistration/azurebot.bicep' = {
      name: 'Azure-Bot-registration'
      params: {
        resourceBaseName: resourceBaseName
        botAadAppClientId: botEntraAppClientId
        botAppDomain: botAppDomain
        botDisplayName: botDisplayName
        graphEntraAppClientId: graphEntraAppClientId
        graphEntraAppClientSecret: graphEntraAppClientSecret
        connectionName: connectionName
      }
    }
    ```

1. Save your changes.

Finally, update the bot registration Bicep file.

1. In the **infra/botRegistration** folder, open the file named **azurebot.bicep**
1. In the file, after **botAppDomain** parameter declaration, add the **graphEntraAppClientId**, **graphEntraAppClientSecret**, and **connectionName** parameter declarations

    ```bicep
    param graphEntraAppClientId string
    @secure()
    param graphEntraAppClientSecret string
    param connectionName string
    ```

1. After the **botServiceM365ExtensionsChannel** resource, add a new resource for the Microsoft Graph connection

    ```bicep
    resource botServicesMicrosoftGraphConnection 'Microsoft.BotService/botServices/connections@2022-09-15' = {
      parent: botService
      name: connectionName
      location: 'global'
      properties: {
        serviceProviderDisplayName: 'Azure Active Directory v2'
        serviceProviderId: '30dd229c-58e3-4a48-bdfd-91ec48eb906c'
        clientId: graphEntraAppClientId
        clientSecret: graphEntraAppClientSecret
        scopes: 'email offline_access openid profile Sites.ReadWrite.All'
        parameters: [
          {
            key: 'tenantID'
            value: 'common'
          }
          {
            key: 'tokenExchangeUrl'
            value: 'api://${botAppDomain}/botid-${ botAadAppClientId}'
          }
        ]
      }
    }
    ```

1. Save your changes

## Task 4 - Authenticate user queries

Next, you add code to authenticate users when they initiate a search using the message extension.

Continuing in Visual Studio:

1. In the **Search** folder, open the file named **SearchApp.cs**
1. In the file, start by adding the **Authentication** namespace from the Bot Framework SDK.

    ```csharp
    using Microsoft.Bot.Connector.Authentication;
    ```

1. In the **OnTeamsMessagingExtensionQueryAsync** method, add the following code at the beginning of the method:

    ```csharp
    var userTokenClient = turnContext.TurnState.Get<UserTokenClient>();
    var tokenResponse = await GetToken(userTokenClient, query.State, turnContext.Activity.From.Id, turnContext.Activity.ChannelId, connectionName, cancellationToken);
    
    if (!HasToken(tokenResponse))
    {
        return await CreateAuthResponse(userTokenClient, connectionName, (Activity)turnContext.Activity, cancellationToken);
    }
    ```

The above code block uses three methods that handle user authentication.

- **GetToken** uses the token service client to obtain an access token for the current user
- **HasToken** checks if the response from the token service contains an access token
- **CreateAuthResponse** called if no token is returned and returns a response, which displays a sign-in link in the user interface

Now, create the methods in the **SearchApp** class.

- Create the **GetToken** method using the following code:

```csharp
private static async Task<TokenResponse> GetToken(UserTokenClient userTokenClient, string state, string userId, string channelId, string connectionName, CancellationToken cancellationToken)
{
  var magicCode = string.Empty;

  if (!string.IsNullOrEmpty(state))
  {
    if (int.TryParse(state, out var parsed)) 
    {
        magicCode = parsed.ToString();
    }
  }

  return await userTokenClient.GetUserTokenAsync(userId, connectionName, channelId, magicCode, cancellationToken);
}
```

First you check to see if the state parameter isn't null or empty, the code attempts to parse it as an integer using the int.TryParse method. If the parsing is successful, the parsed value is assigned to the magicCode variable as a string. The magicCode is then passed as an argument to the GetUserTokenAsync method along with other parameters. The GetUserTokenAsync method uses the magicCode to verify the authenticity of the request. It ensures that the user token is requested by the same entity that initiated the authentication process.

- Create the HasToken method using the following code:

```csharp
private static bool HasToken(TokenResponse tokenResponse)
{
    return tokenResponse != null && !string.IsNullOrEmpty(tokenResponse.Token);
}
```

You check whether a valid token is obtained from the token service by checking if the token response and Token property on the response, isn't empty or null.

- Create the **CreateAuthResponse** method using the following code:

```csharp
private static async Task<MessagingExtensionResponse> CreateAuthResponse(UserTokenClient userTokenClient, string connectionName, Activity activity, CancellationToken cancellationToken)
{
    var resource = await userTokenClient.GetSignInResourceAsync(connectionName, activity, null, cancellationToken);

    return new MessagingExtensionResponse
    {
        ComposeExtension = new MessagingExtensionResult
        {
            Type = "auth",
            SuggestedActions = new MessagingExtensionSuggestedAction
            {
                Actions = new List<CardAction>
                {
                    new() {
                        Type = ActionTypes.OpenUrl,
                        Value = resource.SignInLink,
                        Title = "Sign In",
                    },
                },
            },
        },
    };
}
```

First, you use the **GetSignInResourceAsync** method to retrieve the sign-in link from the token service. The sign-in link is used to construct a **MessagingExtensionResponse** object. You create a new object and set the **ComposeExtension** property of the response to a new **MessagingExtensionResult** object. The type property of the result is set to "auth", indicating that the result is an authentication response. The **SuggestedActions** property of the result is set to a new **MessagingExtensionSuggestedAction** object. The Actions property of the suggested actions is set to a list containing a single **CardAction** object. This **CardAction** object represents an action that can be taken by the user. The Type property of the **CardAction** is set to **ActionTypes.OpenUrl**, indicating that it's an action that opens a URL. The Value property is set to the sign-in link retrieved from the resource. The Title property is set to "Sign In", which specifies the title of the action. Finally, the constructed response is returned from the method.

When a user follows the sign-in link they're taken to a resource hosted on an external domain. The domain must be included in the app manifest file. Add the Bot Framework Token Service domain to the app manifest.

Continuing in Visual Studio:

1. In the **appPackage** folder, open **manifest.json**
1. In the file, update the **validDomains** array, add the domain of the token service:

    ```json
    "validDomains": [
        "token.botframework.com",
        "${{BOT_DOMAIN}}"
    ]    
    ```

1. Save your changes

## Task 5 - Provision resources

With everything now in place, run the Prepare Teams App Dependencies process to provision the required resources.

Continuing in Visual Studio:

1. In **Solution Explorer**, right-click the **MsgExtProductSupport** project
1. Expand the **Teams Toolkit** menu, select **Prepare Teams App Dependencies**
1. In the **Microsoft 365 account** dialog, select **Continue**
1. In the **Provision** dialog, select **Provision**
1. In the **Teams Toolkit warning** dialog, select **Provision**
1. In the **Teams Toolkit information** dialog, select **View provisioned resources** to open a new browser window.

Take a moment to explore the resources that are created and updated in Azure.

## Task 6 - Run and debug

Now start the web service and test the message extension in Microsoft Teams.

Continuing in Visual Studio:

1. Press **F5** to start a debugging session and open a new browser window that navigates to the Microsoft Teams web client.
1. In the browser and if necessary, enter your Microsoft 365 account credentials and continue to Microsoft Teams.
1. In the app install dialog, select **Add**
1. Open a new, or existing Microsoft Teams chat
1. In the message compose area, select **…** to open the app flyout
1. In the list of apps, select **Contoso products** to open the message extension
1. In the text box, enter **Bot Builder** to start a search
1. In the list of results, **select a result** to embed a card into the compose message box
1. A message, **You'll need to sign in to use this app** is shown
1. Select the **sign in link**, to open a new tab and start the authentication flow
1. In the permissions consent page, review the permissions being requested
1. Select *Accept*, to close the tab and return you to Microsoft Teams
1. In the message compose area, select **…** to open the app flyout
1. In the list of apps, select **Contoso products** to open the message extension
1. In the text box, enter **Bot Builder** to start a search
1. You're prompted to sign-in again. Follow the **sign in link** again to start the search.
1. In the list of results, **select a result** to embed a card into the compose message box

Close the browser to stop the debugging session.

[Continue to the next exercise...](./4-exercise-retrieve-product-information-from-sharepoint-online.md)