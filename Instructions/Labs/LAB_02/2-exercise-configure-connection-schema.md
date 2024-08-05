---
lab:
    title: 'Exercise 1 - Configure an external connection and deploy schema'
    module: 'LAB 02: Integrate external content with Copilot for Microsoft 365 using Microsoft Graph connectors built with .NET'
---

# Exercise - Configure an external connection and deploy schema

In this exercise, you build a custom Microsoft Graph connector as a console application. You register a new Microsoft Entra app registration and add the code to create an external connection and deploy its schema.

## Before you start

This exercise will take you about **XX minutes** to complete.

## Task 1 - Register a new Microsoft Entra app registration

Start, by registering a new Entra app registration, which the custom Graph connector uses to authenticate with Microsoft 365.

In a web browser:

1. Go to the **Azure portal** at [https://portal.azure.com](https://portal.azure.com).
1. From the navigation, select **view** underneath **Microsoft Entra ID**.
1. From the side navigation, expand **Manage** and select **App registrations**.
1. In the top navigation, select **New registration**.
1. Specify the following values:
   1. **Name:** MSGraph docs Graph connector
   1. **Supported account types:** Accounts in this organizational directory only (single tenant)
1. Select **Register** to confirm your entry
1. From the overview screen, copy the values of the **Application ID** and **Directory (tenant) ID** properties. You'll need them later.

## Task 2 - Create a credential

Because this custom Graph connector runs without user interaction, you need to configure it to authenticate automatically. For simplicity, create a secret.

Continuing in the web browser:

1. From the side navigation, expand **Manage** and select **Certificates & secrets**.
1. Select the **Client secrets** tab and then select **New client secret**.
1. Enter a **description** of **MSGraph docs Graph connector secret**.
1. Create the secret by selecting **Add**.
1. Copy the **Value** of the newly created secret. You'll need it later.

## Task 3 - Grant API permissions

The final step of configuring the Entra app registration is to grant it API permissions so that it can create the external connection and schema.

Continuing in the web browser:

1. From the side navigation, select **API permissions**.
1. Select **Add a permission**.
1. From the list of APIs, select **Microsoft Graph**.
1. Next, select **Application permissions**.
1. In the filter text box, enter **external**.
1. Expand the **ExternalConnection** section and select the **ExternalConnection.ReadWrite.OwnedBy** permission.
1. Expand the **ExternalItem** section and select the **ExternalItem.ReadWrite.OwnedBy** permission.
1. To confirm your choice, select the **Add permissions** button.
1. To complete the configuration, grant admin consent by selecting the **Grant admin consent for (tenant)** button.
1. Confirm the dialog by selecting **Yes**.

## Task 4 - Create a new console app and install dependencies

After you configured the Entra app registration, the next step is to create a console app, where you'll implement the Graph connector's code.

Open a Windows terminal to create a newe console application:

1. Create a new folder by entering `mkdir documents\console_app` and then navigate to the new folder by entering `cd .\documents\console_app`.
1. Create a new console application by running `dotnet new console`
1. Add dependencies, which you need to build the connector:
   1. To add the library needed to authenticate with Microsoft 365, run `dotnet add package Azure.Identity`.
   1. To add the client library to communicate with Graph APIs, run `dotnet add package Microsoft.Graph`.
   1. To add the library needed to work with user secrets, which you'll configure in the next step, run `dotnet add package Microsoft.Extensions.Configuration.UserSecrets`.
   1. You implement the Graph connector as a console app with two commands: one to create the external connection and deploy the schema, and another to import the content. To support defining commands in your app, run `dotnet add package System.CommandLine --prerelease`.

## Task 5 - Securely store Entra app registration information

After you created the Entra app registration you noted its information such as the application- and tenant ID, and the secret. You use this information in your connector to authenticate with Microsoft 365. To make it available in your code, you store it securely with your project as user secrets.

In a terminal:

1. Ensure that your working directory is set to your newly created console app.
1. To initiate user secrets, run `dotnet user-secrets init`.
1. To securely store the information about the app registration, replace the tokens with the actual values you copied previously and run:

   ```dotnetcli
   dotnet user-secrets set "EntraId:ClientId" "[application id]"
   dotnet user-secrets set "EntraId:ClientSecret" "[secret value]"
   dotnet user-secrets set "EntraId:TenantId" "[directory (tenant) id]"
   ```

## Task 6 - Create Microsoft Graph client

Custom Graph connectors use Microsoft Graph API to manage their external connection and items. Start by creating an instance of the `GraphServiceClient` class from the **Microsoft.Graph** NuGet package you installed in the project.

1. Open your project in Visual Studio 2022.
1. In your project, add a new code file named **GraphService.cs**.
1. In the file, start by adding references to the namespaces you'll use, by adding:

   ```csharp
   using Azure.Identity;
   using Microsoft.Graph;
   using Microsoft.Extensions.Configuration;
   ```

1. Next, define a new class named GraphService:

   ```csharp
   class GraphService
   {
   }
   ```

1. In the `GraphService` class, define a singleton for storing an instance of `GraphServiceClient` to communicate with Microsoft Graph APIs:

   ```csharp
   class GraphService
   {
     static GraphServiceClient? _client;

     public static GraphServiceClient Client
     {
       get
       {
         // TODO: implement
       }
     }
   }
   ```

1. In the `Client` singleton, implement the getter, so that it creates a new instance of `GraphServiceClient` if it doesn't exist yet.

   ```csharp
   public static GraphServiceClient Client
   {
     get
     {
       if (_client is null)
       {
         // TODO: implement
       }
       return _client;
     }
   }
   ```

1. Inside the getting, create a new instance of `GraphServiceClient`, using a credential with the information about the Entra app registration you stored previously:

   ```csharp
   var builder = new ConfigurationBuilder().AddUserSecrets<GraphService>();
   var config = builder.Build();

   var clientId = config["EntraId:ClientId"];
   var clientSecret = config["EntraId:ClientSecret"];
   var tenantId = config["EntraId:TenantId"];

   var credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
   _client = new GraphServiceClient(credential);
   ```

   You start, with creating a configuration builder to access the information about the Entra app registration stored in user secrets. Next, you use the builder to retrieve app registration information. Then, you create a new client secret credential passing the tenant- and client ID, and client secret. Finally, you create an instance of `GraphServiceClient` passing the newly created credential.

1. The complete code looks as follows:

   ```csharp
   using Azure.Identity;
   using Microsoft.Graph;
   using Microsoft.Extensions.Configuration;
   
   class GraphService
   {
     static GraphServiceClient? _client;
   
     public static GraphServiceClient Client
     {
       get
       {
         if (_client is null)
         {
           var builder = new ConfigurationBuilder().AddUserSecrets<GraphService>();
           var config = builder.Build();
     
           var clientId = config["EntraId:ClientId"];
           var clientSecret = config["EntraId:ClientSecret"];
           var tenantId = config["EntraId:TenantId"];
           
           var credential = new ClientSecretCredential(tenantId, clientId,    clientSecret);
           _client = new GraphServiceClient(credential);
         }
   
         return _client;
       }
     }
   }
   ```

1. Save your changes

## Task 7 - Define external connection and schema configuration

The next step is to define the external connection and schema that the Graph connector should use. Because the connector's code needs access to the external connection's ID in several places, store it in a central place in your code.

In the code editor:

1. Create a new file named **ConnectionConfiguration.cs**.
1. Add a reference to the namespace with Microsoft Graph models:

   ```csharp
   using Microsoft.Graph.Models.ExternalConnectors;
   ```

1. Next, in the same file, define a new static class named `ConnectionConfiguration`:

   ```csharp
   static class ConnectionConfiguration
   {
   
   }
   ```

1. In the `ConnectionConfiguration` class, add a new property named `ExternalConnection`. Implement it to return an instance of the `ExternalConnection` Microsoft Graph Model:

   ```csharp
   public static ExternalConnection ExternalConnection
   {
     get
     {
       return new ExternalConnection
       {
         Id = "msgraphdocs",
         Name = "Microsoft Graph documentation",
         Description = "Documentation for Microsoft Graph API which explains what Microsoft Graph is and how to use it."
       };
     }
   }
   ```

1. Next, add another property named `Schema`. Implement it to return an instance of the `Schema` Graph Model:

   ```csharp
   public static Schema Schema
   {
     get
     {
       return new Schema
       {
         BaseType = "microsoft.graph.externalItem",
         Properties = new()
         {
           // TODO: implement
         }
       };
     }
   }
   ```

1. In the schema, define properties that you track for every external item that you ingest using the connector:

   ```csharp
   new Property
   {
     Name = "title",
     Type = PropertyType.String,
     IsQueryable = true,
     IsSearchable = true,
     IsRetrievable = true,
     Labels = new() { Label.Title }
   },
   new Property
   {
     Name = "description",
     Type = PropertyType.String,
     IsQueryable = true,
     IsSearchable = true,
     IsRetrievable = true
   },
   new Property
   {
     Name = "iconUrl",
     Type = PropertyType.String,
     IsRetrievable = true,
     Labels = new() { Label.IconUrl }
   },
   new Property
   {
     Name = "url",
     Type = PropertyType.String,
     IsRetrievable = true,
     Labels = new() { Label.Url }
   }
   ```

   You start, with defining the **title** property, which stores the title of the external item imported to Microsoft 365. The item's title is a part of the full-text index (`IsSearchable = true`). Users can also explicitly query for its contents in keyword queries (`IsQueryable = true`). The title can be also retrieved and displayed in search results (`IsRetrievable = true`). The **title** property represents the item's title, which you indicate using the `Title` semantic label.

   Next, you define the **description** property, which stores the summary of the contents of the external item. Its definition is similar to the title. There's however no semantic label for the description, which is why you don't define it.

   Next, you define a property to store the URL of the icon for each item. Copilot for Microsoft 365 requires this property and it needs to be mapped using the `IconUrl` semantic label.

   Finally, you define the **url** property, which stores the original URL of the external item. Users use this URL to navigate to the external item from search results or Copilot from Microsoft 365. URL is one of the properties that Copilot for Microsoft 365 requires, which is why you map it using the `Url` semantic label.

1. The complete code looks as follows:

   ```csharp
   using Microsoft.Graph.Models.ExternalConnectors;
   static class ConnectionConfiguration
   {
     public static ExternalConnection ExternalConnection
     {
       get
       {
         return new ExternalConnection
         {
           Id = "msgraphdocs",
           Name = "Microsoft Graph documentation",
           Description = "Documentation for Microsoft Graph API which    explains what Microsoft Graph is and how to use it."
         };
       }
     }
     public static Schema Schema
     {
       get
       {
         return new Schema
         {
           BaseType = "microsoft.graph.externalItem",
           Properties = new()
           {
             new Property
             {
               Name = "title",
               Type = PropertyType.String,
               IsQueryable = true,
               IsSearchable = true,
               IsRetrievable = true,
               Labels = new() { Label.Title }
             },
             new Property
             {
               Name = "description",
               Type = PropertyType.String,
               IsQueryable = true,
               IsSearchable = true,
               IsRetrievable = true
             },
             new Property
             {
               Name = "iconUrl",
               Type = PropertyType.String,
               IsRetrievable = true,
               Labels = new() { Label.IconUrl }
             },
             new Property
             {
               Name = "url",
               Type = PropertyType.String,
               IsRetrievable = true,
               Labels = new() { Label.Url }
             }
           }
         };
       }
     }
   }
   ```

1. Save your changes

## Task 8 - Create external connection

Continue with adding code that uses the information about the external connection you defined in the previous section to create the external connection in Microsoft 365.

In the code editor:

1. Create a new file named **ConnectionService.cs**.
1. In the file, start by adding references to the namespaces:

   ```csharp
   using Microsoft.Graph.Models.ExternalConnectors;
   ```

1. Next, in the same file, define a new static class named `ConnectionService`:

   ```csharp
   static class ConnectionService
   {
   
   }
   ```

1. In the `ConnectionService` class, add a new method named `CreateConnection`:

   ```csharp
   async static Task CreateConnection()
   {

   }
   ```

1. In the `CreateConnection` method, use the Microsoft Graph client instance to call Microsoft Graph APIs and create the external connection using the previously defined connection information:

   ```csharp
   Console.Write("Creating connection...");
   
   await GraphService.Client.External.Connections
     .PostAsync(ConnectionConfiguration.ExternalConnection);
   
   Console.WriteLine("DONE");
   ```

1. The complete code looks as follows:

   ```csharp
   using Microsoft.Graph.Models.ExternalConnectors;
   static class ConnectionService
   {
     async static Task CreateConnection()
     {
       Console.Write("Creating connection...");
   
       await GraphService.Client.External.Connections
         .PostAsync(ConnectionConfiguration.ExternalConnection);
   
       Console.WriteLine("DONE");
     }
   }
   ```

1. Save your changes.

Before you test this code, add the code to create the schema. That way, you can test the complete flow of creating and configuring the external connection.

## Task 9 - Create external connection schema

The last part of creating an external connection is to create its schema.

In the code editor:

1. Open the **ConnectionService.cs** file.
1. In the ConnectionService class, add a new method named `CreateSchema`:

   ```csharp
   async static Task CreateSchema()
   {
   }
   ```

1. In the `CreateSchema` method, use the Microsoft Graph client instance to call the Microsoft Graph API to create the schema. Then, wait for its creation.

   ```csharp
   Console.WriteLine("Creating schema...");
   
   await GraphService.Client.External
     .Connections[ConnectionConfiguration.ExternalConnection.Id]
     .Schema
     .PatchAsync(ConnectionConfiguration.Schema);
   
   do
   {
     var externalConnection = await GraphService.Client.External
       .Connections[ConnectionConfiguration.ExternalConnection.Id]
       .GetAsync();
   
     Console.Write($"State: {externalConnection?.State.ToString()}");
   
     if (externalConnection?.State != ConnectionState.Draft)
     {
       Console.WriteLine();
       break;
     }
   
     Console.WriteLine($". Waiting 60s...");
   
     await Task.Delay(60_000);
   }
   while (true);
   
   Console.WriteLine("DONE");
   ```

1. In the same file, add a new method named `ProvisionConnection`. In its code, call the `CreateConnection` and `CreateSchema` methods you defined previously:

   ```csharp
   public static async Task ProvisionConnection()
   {
     try
     {
       await CreateConnection();
       await CreateSchema();
     }
     catch (Exception ex)
     {
       Console.WriteLine(ex.Message);
     }
   }
   ```

1. The complete code looks as follows:

   ```csharp
   using Microsoft.Graph.Models.ExternalConnectors;
   
   static class ConnectionService
   {
     async static Task CreateConnection()
     {
       Console.Write("Creating connection...");
   
       await GraphService.Client.External.Connections
         .PostAsync(ConnectionConfiguration.ExternalConnection);
   
       Console.WriteLine("DONE");
     }
   
     async static Task CreateSchema()
     {
       Console.WriteLine("Creating schema...");
   
       await GraphService.Client.External
         .Connections[ConnectionConfiguration.ExternalConnection.Id]
         .Schema
         .PatchAsync(ConnectionConfiguration.Schema);
   
       do
       {
         var externalConnection = await GraphService.Client.External
           .Connections[ConnectionConfiguration.ExternalConnection.Id]
           .GetAsync();
   
         Console.Write($"State: {externalConnection?.State.ToString()}");
   
         if (externalConnection?.State != ConnectionState.Draft)
         {
           Console.WriteLine();
           break;
         }
   
         Console.WriteLine($". Waiting 60s...");
   
         await Task.Delay(60_000);
       }
       while (true);
   
       Console.WriteLine("DONE");
     }
   
     public static async Task ProvisionConnection()
     {
       try
       {
         await CreateConnection();
         await CreateSchema();
       }
       catch (Exception ex)
       {
         Console.WriteLine(ex.Message);
       }
     }
   }
   ```

1. Save your changes.

## Task 10 - Test the code

The last step left is to create an entry point in the application, which will create the connection and its schema. You do this by creating a command that you invoke by starting the application from the command line.

In the code editor:

1. Open the **Program.cs** file.
1. Replace its contents with the following code:

   ```csharp
   using System.CommandLine;
   
   var provisionConnectionCommand = new Command("provision-connection",    "Provisions external connection");
   provisionConnectionCommand.SetHandler(ConnectionService.   ProvisionConnection);
   
   var rootCommand = new RootCommand();
   rootCommand.AddCommand(provisionConnectionCommand);
   Environment.Exit(await rootCommand.InvokeAsync(args));
   ```

   You start by defining a command named `provision-connection.` This command invokes the `ConnectionService.ProvisionConnection` method that you defined previously. Finally, you register the command with the command-line processor and start the application monitoring arguments passed in command line.

1. Save your changes

To test the application:

1. Open a terminal.
1. Change the working directory to the project folder.
1. Run `dotnet build` to build the project.
1. Start the app by running `dotnet run -- provision-connection`.
1. Wait several minutes for the connection and schema to be created.

[Continue to the next exercise...](./3-exercise-import-external-content.md)