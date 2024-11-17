---
lab:
    title: 'Exercise 4 - Explore the plugin source code'
    module: 'LAB 02: Build your own message extension plugin with TypeScript (TS) for Microsoft Copilot'
---

# Exercise 4 - Explore the plugin source code

In this exercise, you'll review the application code so that you can understand how a **Message Extension** works.

## Task 1 - Examine the manifest

The core of any Microsoft 365 application is its application manifest. This is where you provide the information that Microsoft 365 needs to access your application.

In your **working directory**, open the **appPackackage/manifest.json** file. This JSON file is placed in a zip archive with two icon files to create the application package. The **icons** property includes paths to these icons.

```json
"icons": {
    "color": "Northwind-Logo3-192-${{TEAMSFX_ENV}}.png",
    "outline": "Northwind-Logo3-32.png"
},
```

Notice the token `${{TEAMSFX_ENV}}` in one of the icon names. Teams Toolkit will replace this token with your environment name, such as **local** or **dev** (for an Azure deployment in development). Thus, the icon color will change depending on the environment.

### Application description

Now have a look at the **name** and **description**. Notice that the **description** is quite long! This is important so both users and Copilot can learn what your application does and when to use it.

```json
    "name": {
        "short": "Northwind Inventory",
        "full": "Northwind Inventory App"
    },
    "description": {
        "short": "App allows you to find and update product inventory information",
        "full": "Northwind Inventory is the ultimate tool for managing your product inventory. With its intuitive interface and powerful features, you'll be able to easily find your products by name, category, inventory status, and supplier city. You can also update inventory information with the app. \n\n **Why Choose Northwind Inventory:** \n\n Northwind Inventory is the perfect solution for businesses of all sizes that need to keep track of their inventory. Whether you're a small business owner or a large corporation, Northwind Inventory can help you stay on top of your inventory management needs. \n\n **Features and Benefits:** \n\n - Easy Product Search through Microsoft Copilot. Simply start by saying, 'Find northwind dairy products that are low on stock' \r - Real-Time Inventory Updates: Keep track of inventory levels in real-time and update them as needed \r  - User-Friendly Interface: Northwind Inventory's intuitive interface makes it easy to navigate and use \n\n **Availability:** \n\n To use Northwind Inventory, you'll need an active Microsoft 365 account . Ensure that your administrator enables the app for your Microsoft 365 account."
    },
```

### Bot definition

Scroll down a bit to **composeExtensions**. Compose extension is the historical term for message extension; this is where the app's message extensions are defined. Message extensions communicate using the Azure Bot Framework; this provides a fast and secure communication channel between Microsoft 365 and your application. When you first ran your project, Teams Toolkit registered a bot, and will place its **botID** here.

```json
    "composeExtensions": [
        {
            "botId": "${{BOT_ID}}",
            "commands": [
                {
                    ...
```

### Command definitions

This message extension has two commands, which are defined in the `commands` array. If you completed the previous exercise, there will be also a third command to search by company name. Let's skip the first command for a moment since it's the most complex one. The following command allows Copilot (or a user) to search for discounted products within a Northwind category. This command accepts a single parameter, **categoryName**.

```json
{
    "id": "discountSearch",
    "context": [
        "compose",
        "commandBox"
    ],
    "description": "Search for discounted products by category",
    "title": "Discounts",
    "type": "query",
    "parameters": [
        {
            "name": "categoryName",
            "title": "Category name",
            "description": "Enter the category to find discounted products",
            "inputType": "text"
        }
    ]
},
```

Now let's move back to the first command, **inventorySearch**, which has 5 parameters, allowing for much more sophisticated queries.

```json
{
    "id": "inventorySearch",
    "context": [
        "compose",
        "commandBox"
    ],
    "description": "Search products by name, category, inventory status, supplier location, stock level",
    "title": "Product inventory",
    "type": "query",
    "parameters": [
        {
            "name": "productName",
            "title": "Product name",
            "description": "Enter a product name here",
            "inputType": "text"
        },
        {
            "name": "categoryName",
            "title": "Category name",
            "description": "Enter the category of the product",
            "inputType": "text"
        },
        {
            "name": "inventoryStatus",
            "title": "Inventory status",
            "description": "Enter what status of the product inventory. Possible values are 'in stock', 'low stock', 'on order', or 'out of stock'",
            "inputType": "text"
        },
        {
            "name": "supplierCity",
            "title": "Supplier city",
            "description": "Enter the supplier city of product",
            "inputType": "text"
        },
        {
            "name": "stockQuery",
            "title": "Stock level",
            "description": "Enter a range of integers such as 0-42 or 100- (for >100 items). Only use if you need an exact numeric range.",
            "inputType": "text"
        }
    ]
},
```

Copilot is able to fill these in based on the descriptions and the message extension will return a list of products filtered by all the nonblank parameters.

## Task 2 - Examine the Bot code

Now open the  **src/searchApp.ts** file, which contains the code for the bot that uses the [Bot Builder SDK](https://learn.microsoft.com/azure/bot-service/index-bf-sdk) to communicate with the Azure Bot Framework. Notice that the bot extends an SDK class **TeamsActivityHandler**.

```typescript
export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  ...
```

### Message extension query

The application is able to handle messages (called **activities**) coming from Microsoft 365 by overriding the methods of **TeamsActivityHandler**.

The first of these is a **Messaging Extension Query** activity. This function is called when a user types into a message extension or when Copilot calls it. The handler is dispatching the query based on the **commandID**. These are the same commandID's used in the app manifest.

```typescript
  // Handle search message extension
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {

    switch (query.commandId) {
      case productSearchCommand.COMMAND_ID: {
        return productSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
      }
      case discountedSearchCommand.COMMAND_ID: {
        return discountedSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
      }
    }
  }
  ...
```

### Adaptive card actions

The other type of activities our app needs to handle are the adaptive card actions, such as when a user selects **Update stock** or **Reorder** on an adaptive card. Since there's no specific method for an adaptive card action, the code overrides `onInvokeActivity()`, which is a much broader class of activity that includes message extension queries. For that reason, the code manually checks the activity name, and dispatches to the appropriate handler. If the activity name isn't for an adaptive card action, the `else` clause runs the base implementation of `onInvokeActivity()`, which among other things, will call our `handleTeamsMessagingExtensionQuery()` method if the **Invoke** activity is a query.

```typescript
import {
  TeamsActivityHandler,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
  InvokeResponse
} from "botbuilder";
import productSearchCommand from "./messageExtensions/productSearchCommand";
import discountedSearchCommand from "./messageExtensions/discountSearchCommand";
import revenueSearchCommand from "./messageExtensions/revenueSearchCommand";
import actionHandler from "./adaptiveCards/cardHandler";

export class SearchApp extends TeamsActivityHandler {
  constructor() {
    super();
  }

  // Handle search message extension
  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<MessagingExtensionResponse> {

    switch (query.commandId) {
      case productSearchCommand.COMMAND_ID: {
        return productSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
      }
      case discountedSearchCommand.COMMAND_ID: {
        return discountedSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
      }
    }

  }

  // Handle adaptive card actions
  public async onInvokeActivity(context: TurnContext): Promise<InvokeResponse> {
    let runEvents = true;
    // console.log (`🎬 Invoke activity received: ${context.activity.name}`);
    try {
      if(context.activity.name==='adaptiveCard/action'){
        switch (context.activity.value.action.verb) {
          case 'ok': {
            return actionHandler.handleTeamsCardActionUpdateStock(context);
          }
          case 'restock': {
            return actionHandler.handleTeamsCardActionRestock(context);
          }
          case 'cancel': {
            return actionHandler.handleTeamsCardActionCancelRestock(context);
          }
          default:
            runEvents = false;
            return super.onInvokeActivity(context);
        }
      } else {
          runEvents = false;
          return super.onInvokeActivity(context);
      }
    } 
```

## Task 3 - Examine the message extension command code

In an effort to make the code more modular, readable, and reusable, each message extension command is placed in its own TypeScript module. Have a look at **src/messageExtensions/discountSearchCommand.ts** for an example.

First, observe that the module exports a constant `COMMAND_ID`, which contains the same **commandID** found in the app manifest, and allows the switch statement in **searchApp.ts** to work properly.

Then it provides a function, `handleTeamsMessagingExtensionQuery()`, to handle incoming queries for **discounted products by category**.

```typescript
async function handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
): Promise<MessagingExtensionResponse> {

    // Seek the parameter by name, don't assume it's in element 0 of the array
    let categoryName = cleanupParam(query.parameters.find((element) => element.name === "categoryName")?.value);
    console.log(`💰 Discount query #${++queryCount}: Discounted products with categoryName=${categoryName}`);

    const products = await getDiscountedProductsByCategory(categoryName);

    console.log(`Found ${products.length} products in the Northwind database`)
    const attachments = [];
    products.forEach((product) => {
        const preview = CardFactory.heroCard(product.ProductName,
            `Avg discount ${product.AverageDiscount}%<br />Supplied by ${product.SupplierName} of ${product.SupplierCity}`,
            [product.ImageUrl]);

        const resultCard = cardHandler.getEditCard(product);
        const attachment = { ...resultCard, preview };
        attachments.push(attachment);
    });
    return {
        composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments: attachments,
        },
    };
}
```

Notice that the index in the `query.parameters` array might not correspond to the parameter's position in the manifest. While this is generally only an issue for a multi-parameter command, the code will still get the value based on the parameter name rather than hard coding an index.

After cleaning up the parameter (trimming it, and handling the fact that sometimes Copilot assumes "**\***" is a wildcard that matches everything), the code calls the Northwind data access layer to `getDiscountedProductsByCategory()`.

Then it iterates through the products and creates two cards for each:

- a _preview_ card, which is implemented as a **hero** card. This is displayed in the search results in the user interface and in some citations in Copilot.

- a _result_ card, which is implemented as an **adaptive** card that includes all the details.

In the next task, we'll review the adaptive card code and check out the Adaptive Card designer.

## Task 4 - Examine the adaptive cards and related code

The project's adaptive cards are in the **src/adaptiveCards/** folder. There are 3 cards, each implemented as a JSON file.

- **editCard.json** - This is the initial card displayed by the message extension or a Copilot reference.

- **successCard.json** - When a user takes action, this card is displayed to indicate success. It's mostly the same as the edit card except that it includes a message to the user.

- **errorCard.json** - If an action fails, this card is displayed.

Let's take a look at the edit card in the **Adaptive Card Designer**. Open your web browser to [https://adaptivecards.io](https://adaptivecards.io) and select the **Designer** option at the top.

![Screenshot of the Adaptive Card Designer.](../media/5-01-adaptive-card-designer-01.png)

Notice the data binding expressions such as `"text": "📦 ${productName}",`. This binds the `productName` property in the data to the text on the card.

Now select **Microsoft Teams** as the host application 1️⃣. Paste the entire contents of **editCard.json** into the Card Payload Editor 2️⃣, and the contents of **sampleData.json** into the Sample Data Editor 3️⃣. The sample data is identical to a product as provided in the code. You should see the card as rendered, except for a small error that arises due to the designer's inability to display one of the adaptive card formats.

![Screenshot of Copilot the adaptive card rendered based on the json.](../media/5-01-adaptive-card-designer-02.png)

Near the top of the page, try changing the **Theme** and **Emulated Device** to see how the card would look in dark theme or on a mobile device. This is the tool that was used to build adaptive cards for the sample application.

Now, back in Visual Studio Code, open **cardHandler.ts**. The function `getEditCard()` is called from each of the message extension commands to obtain a **result** card. The code reads the adaptive card JSON - which is considered a template - and then binds it to product data. The result is more JSON - the same card as the template, with the data binding expressions all filled in. Finally, the `CardFactory` module is used to convert the final JSON into an adaptive card object for rendering.

```typescript
function getEditCard(product: ProductEx): any {

    var template = new ACData.Template(editCard);
    var card = template.expand({
        $root: {
            productName: product.ProductName,
            unitsInStock: product.UnitsInStock,
            productId: product.ProductID,
            categoryId: product.CategoryID,
            imageUrl: product.ImageUrl,
            supplierName: product.SupplierName,
            supplierCity: product.SupplierCity,
            categoryName: product.CategoryName,
            inventoryStatus: product.InventoryStatus,
            unitPrice: product.UnitPrice,
            quantityPerUnit: product.QuantityPerUnit,
            unitsOnOrder: product.UnitsOnOrder,
            reorderLevel: product.ReorderLevel,
            unitSales: product.UnitSales,
            inventoryValue: product.InventoryValue,
            revenue: product.Revenue,
            averageDiscount: product.AverageDiscount
        }
    });
    return CardFactory.adaptiveCard(card);
}
```

Scrolling down, you'll see the handler for each of the action buttons on the card. The card submits data when an action button is clicked - specifically `data.txtStock`, which is the **Quantity** input box on the card, and `data.productId`, which is sent in each card action to let the code know what product to update.

```typescript
async function handleTeamsCardActionUpdateStock(context: TurnContext) {

    const request = context.activity.value;
    const data = request.action.data;
    console.log(`🎬 Handling update stock action, quantity=${data.txtStock}`);

    if (data.txtStock && data.productId) {

        const product = await getProductEx(data.productId);
        product.UnitsInStock = Number(data.txtStock);
        await updateProduct(product);

        var template = new ACData.Template(successCard);
        var card = template.expand({
            $root: {
                productName: product.ProductName,
                unitsInStock: product.UnitsInStock,
                productId: product.ProductID,
                categoryId: product.CategoryID,
                imageUrl: product.ImageUrl,
                ...
```

As you can see, the code obtains these two values, updates the database, and then sends a new card that contains a message and the updated data.

## Congratulations

You have completed Exercise 5 and the Copilot for Microsoft 365 Message Extensions plugin lab. Thanks very much for doing these labs!

[Continue to the lab summary...](./7-summary.md)
