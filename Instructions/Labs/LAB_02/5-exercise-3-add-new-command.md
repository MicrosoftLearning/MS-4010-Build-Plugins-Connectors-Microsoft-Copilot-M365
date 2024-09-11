---
lab:
    title: 'Exercise 3 - Add a new command'
    module: 'LAB 02: Build your own message extension plugin with TypeScript (TS) for Microsoft 365 Copilot'
---

# Exercise 3 - Add a new command

In this exercise, you'll enhance the Teams Message Extension and Copilot plugin by adding a new command. While the current message extension effectively provides information about products within the Northwind inventory database, it doesn't provide information related to Northwind’s customers. You'll introduce a new command associated with an API call that retrieves products ordered by a customer name specified by the user. This exercise assumes you have completed at least exercises 1, 2, and 3. It's fine to skip Exercise 4 in case you don't have a Microsoft 365 Copilot license.

We'll go through the following tasks to accomplish this:

1. **Extend the Message Extension / plugin User Interface** by modifying the Teams app manifest. This includes introducing a new command: **"companySearch"**. Observe that the UI for the Message Extension is an adaptive card, whereas for Copilot it's text input and output in Copilot chat.

1. **Create a handler for the 'companySearch' command**. This will parse the query string passed in from the message routing code, validate the input, and call the product search by company API. This task will also populate an adaptive card with the returned product list, which will be displayed in the message or Copilot chat UI.

1. Update the command **routing** code to route the new command to the handler created in the previous task. You'll do this by extending the method called by the Bot Framework when users query the Northwind database (**handleTeamsMessagingExtensionQuery**). 

1. **Implement Product Search by Company** that returns a list of products ordered by that company.

1. **Run the app** and search of products that were purchased by a specified company.

## Task 1 - Extend the Message Extension / plugin User Interface 

1. In Visual Studio Code from the **working folder**, open **manifest.json** and add the following json immediately after the `discountSearch` command (**line 98**). With this additional information, you're adding to the `commands` array that defines the list of commands supported by the plugin.

   ```json
   {
       "id": "companySearch",
       "context": [
           "compose",
           "commandBox"
       ],
       "description": "Given a company name, search for products ordered by that company",
       "title": "Customer",
       "type": "query",
       "parameters": [
           {
               "name": "companyName",
               "title": "Company name",
               "description": "The company name to find products ordered by that company",
               "inputType": "text"
           }
       ]
   }
   ```

> [!NOTE] 
> The **id** is the connection between the UI and the code. This value is defined as **COMMAND_ID** in the **discount\product\SearchCommand.ts** files. See how each of these files has a unique **COMMAND_ID** that corresponds to the value of **id**.

## Task 2 - Create a handler for the 'companySearch' command

In this exercise, we'll copy some of the existing code to create new handlers for our commands. 

1. In Visual Studio Code under your **working directory**, navigate to **.\src\messageExtensions** and copy '**productSearchCommand.ts**' and paste into the same folder to create a copy. Rename this file **customerSearchCommand.ts**.

1. Change line 7 to:

    ```typescript
    import { searchProductsByCustomer } from "../northwindDB/products";
    ```

1. Change line 10 to:

   ```javascript
   const COMMAND_ID = "companySearch";
   ```



1. Replace the content of **handleTeamsMessagingExtensionQuery** with:

   ```javascript
    {
       let companyName;
   
       // Validate the incoming query, making sure it's the 'companySearch' command
       // The value of the 'companyName' parameter is the company name to search for
       if (query.parameters.length === 1 && query.parameters[0]?.name === "companyName") {
           [companyName] = (query.parameters[0]?.value.split(','));
       } else { 
           companyName = cleanupParam(query.parameters.find((element) => element.name === "companyName")?.value);
       }
       console.log(`🍽️ Query #${++queryCount}:\ncompanyName=${companyName}`);    
   
       const products = await searchProductsByCustomer(companyName);
   
       console.log(`Found ${products.length} products in the Northwind database`)
       const attachments = [];
       products.forEach((product) => {
           const preview = CardFactory.heroCard(product.ProductName,
               `Customer: ${companyName}`, [product.ImageUrl]);
   
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

> [!NOTE]
> You will implement `searchProductsByCustomer` in Task 4.

## Task 3 - Update the command routing

In this task, you'll route the `companySearch` command to the handler you implemented in the previous task.

1. Open **searchApp.ts** and look for this on **line 10**:

   ```javascript
   import discountedSearchCommand from "./messageExtensions/discountSearchCommand";
   ```

1. Add this as **line 11**:

   ```javascript
   import customerSearchCommand from "./messageExtensions/customerSearchCommand";
   ```

1. Underneath this statement:

   ```javascript
         case discountedSearchCommand.COMMAND_ID: {
           return discountedSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
         }
   ```

1. Add this statement:

   ```javascript
         case customerSearchCommand.COMMAND_ID: {
           return customerSearchCommand.handleTeamsMessagingExtensionQuery(context, query);
         }
   ```

> [!NOTE]
> The UI-based operation of the plugin is explicitly called. However, when invoked by Microsoft 365 Copilot, the command is triggered by the Copilot orchestrator.

## Task 4 - Implement Product Search by Company

In this task, you'll implement a product search **by Company name** and return a list of the company's ordered products. The table output from the query will look like this:

| Table         | Find        | Look Up By    |
| ------------- | ----------- | ------------- |
| Customer      | Customer ID | Customer Name |
| Orders        | Order ID    | Customer ID   |
| OrderDetail   | Product     | Order ID      |

Here's how the query works: 

1. Use the **Customer** table to find the **Customer Id** with the **Customer Name**. 

1. Query the **Orders** table with the **Customer Id** to retrieve the associated **Order Ids**. 

1. For each **Order Id**, find the associated products in the **OrderDetail** table. 

1. Finally, return a list of products ordered by the specified company name.

Now, lets modify the **products.ts** file to add the new search query.

1. Open **.\src\northwindDB\products.ts**

1. Update the `import` statement on line 1 to include OrderDetail, Order, and Customer. It should look as follows:

   ```javascript
   import {
       TABLE_NAME, Product, ProductEx, Supplier, Category, OrderDetail,
       Order, Customer
   } from './model';
   ```

1. Underneath this on **line 8**:

   ```javascript
   import { getInventoryStatus } from '../adaptiveCards/utils';
   ```

       Add the new function `searchProductsByCustomer()`:

   ```javascript
   export async function searchProductsByCustomer(companyName: string): Promise<ProductEx[]> {

       let result = await getAllProductsEx();
   
       let customers = await loadReferenceData<Customer>(TABLE_NAME.CUSTOMER);
       let customerId="";
       for (const c in customers) {
           if (customers[c].CompanyName.toLowerCase().includes(companyName.toLowerCase())) {
               customerId = customers[c].CustomerID;
               break;
           }
       }
       
       if (customerId === "") 
           return [];

       let orders = await loadReferenceData<Order>(TABLE_NAME.ORDER);
       let orderdetails = await loadReferenceData<OrderDetail>(TABLE_NAME.ORDER_DETAIL);
       // build an array orders by customer id
       let customerOrders = [];
       for (const o in orders) {
           if (customerId === orders[o].CustomerID) {
               customerOrders.push(orders[o]);
           }
       }
       
       let customerOrdersDetails = [];
       // build an array order details customerOrders array
       for (const od in orderdetails) {
           for (const co in customerOrders) {
               if (customerOrders[co].OrderID === orderdetails[od].OrderID) {
                   customerOrdersDetails.push(orderdetails[od]);
               }
           }
       }

       // Filter products by the ProductID in the customerOrdersDetails array
       result = result.filter(product => 
           customerOrdersDetails.some(order => order.ProductID === product.ProductID)
       );

       return result;
   }
   ```

## Task 5 - Run the App! Search for product by company name

Now you're ready to test the sample as a plugin for Microsoft 365 Copilot.

1. Remove the **Northwest Inventory** app in Teams. This task is necessary since you're updating the manifest. Manifest updates require the app to be reinstalled. The cleanest way to do this is to first remove it from Teams.

    1. In the Teams sidebar, select on the three dots (...) 1️⃣. You should see **Northwind Inventory** 2️⃣ in the list of applications.

    1. Right-click on the **Northwest Inventory** icon and select uninstall 3️⃣.

        ![Screenshot of how to uninstall Northwind Inventory.](../media/3-01-uninstall-app.png)

1. Start the app in Visual Studio Code using the **Debug in Teams (Edge)** profile by pressing **F5**.

1. In Teams, select **Chat** and then **Copilot**. Copilot should be the top-most option.

1. Select the **Plugin icon** and select **Northwind Inventory** to enable the plugin.

1. Enter the prompt: 

   ```console
   What are the products ordered by 'Consolidated Holdings' in Northwind Inventory?
   ```

   The Terminal output shows Copilot understood the query and executed the `companySearch` command, passing company name extracted by Copilot.

   ![Screenshot of Copilot understanding the query and executed the `companySearch` command.](../media/3-08-terminal-query-output.png)

   Here's the output in Copilot:

   ![Screenshot of Copilot producing the results from the command.](../media/3-07-response-customer-search.png)

Here are other prompts to try:

```console
What are the products ordered by 'Consolidated Holdings' in Northwind Inventory? Please list the product name, price and supplier in a table.
```

Of course, you can also test this new command using the sample as a Message Extension, like we did in previous exercise. 

1. In the Teams sidebar, move to the **Chats** section and pick any chat or start a new chat with a colleague.

1. Select the **+** sign to access to the Apps menu.

1. Pick the **Northwind Inventory** app.

1. Notice how now you can see a new tab called **Customer**.

1. Search for **Consolidated Holdings** and see the products ordered by this company. They'll match the ones that Copilot returned you in the previous task.

    ![Screenshot of the new command used as a message extension.](../media/3-08-customer-message-extension.png)

## Check your work

After completing the exercise, you should have a new command to search for orders by company in the Northwind Inventory app. You should also be able to successfully use the plugin with Copilot and as a message extension in other apps. 

In the next exercise, you'll explore the plugin source code and adaptive cards to learn more about how the apps are built and how to customize them further.

[Continue to the next exercise...](./6-exercise-4-explore-plugin-source-code.md)