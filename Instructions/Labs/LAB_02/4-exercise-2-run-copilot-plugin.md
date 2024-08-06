---
lab:
    title: 'Exercise 2 - Run the sample as a Copilot plugin'
    module: 'LAB 03: Build your own message extension plugin with TypeScript (TS) for Microsoft Copilot'
---

# Exercise 2 - Run the sample as a Copilot plugin

In this exercise, you'll run the application as a plugin for Microsoft Copilot for Microsoft 365. You'll experiment with various prompts and you'll observe how the plugin gets invoked using different parameters.

> [!NOTE]  
> To perform the following exercise, your account must have a valid license for Copilot for Microsoft 365.

## Task 1 - Test in Microsoft Copilot for Microsoft 365 (single parameter)

1. In the application rail on the left, select the **Copilot** app.

1. On the right side of the compose box, select the **plugin** icon 1️⃣, and enable the **Northwind Inventory** plugin 2️⃣.

    ![Screenshot of small panel with a toggle for each plugin.](../media/3-02-plugin-panel.png)

1. For best results, select the **New chat** icon at the top right before each prompt or set of related prompts.

    ![Screenshot of Copilot showing its new chat screen.](../media/3-01-new-chat.png)

1. Try the following prompts that use only a single parameter of the message extension:

    - _Find information about Chai in Northwind Inventory._

    - _Find discounted seafood in Northwind. Show a table with the products, supplier names, average discount rate, and revenue per period._

The last one should reference the documents you uploaded to OneDrive. As you're testing, watch the log messages within Visual Studio Code. You should be able to see when Copilot calls your plugin and submits a query. For example, after requesting **discounted seafood items**, Copilot issued this query using the `discountSearch` command.

![Screenshot of log file shows a discount search for seafood.](../media/3-02-a-query-log-1.png)

You might see citations of the Northwind data in 3 forms. If there's a single reference, Copilot might show the whole card.

![Screenshot of adaptive card for Chai embedded in a Copilot response.](../media/3-03-a-response-on-chai.png)

If there are multiple references, Copilot might show a small number next to each. You can hover over these numbers to display the adaptive card. References will also be listed below the response.

![Screenshot of reference numbers embedded in a Copilot response - hovering over the number shows the adaptive card.](../media/3-03-response-on-chai.png)

Try these adaptive cards to take action on the products. Notice that this doesn't affect earlier responses from Copilot.

Feel free to try making up your own prompts. You'll find that they only work if Copilot is able to query the plugin for the required information. This underscores the need to anticipate the kinds of prompts users will issue, and providing corresponding types of queries for each one. Having multiple parameters will make this more powerful!

## Task 2 - Test in Microsoft Copilot for Microsoft 365 (multiple parameters)

In this exercise, you'll try some prompts that exercise the multi-parameter feature in the sample plugin. These prompts will request data that can be retrieved by **name**, **category**, **inventory status**, **supplier city**, and **stock level**, as defined in the **app manifest**.

For example, try prompting **_Find Northwind beverages with more than 100 items in stock_**. To generate its response, Copilot must identify products:

- Where the category is **beverages**.
  
  _AND_

- Where inventory status is **in stock**.

  _AND_

- Where the **stock level** is greater than **100**.

If you look at the log file, you can see that Copilot was able to understand this requirement and fill in 3 of the parameters in the first message extension command.

![Screenshot of log showing a query for categoryName=beverages and stockLevel=100- .](../media/3-06-find-northwind-beverages-with-more-than-100.png)

The plugin code applies all three filters, providing a result set of just 4 products. Copilot uses the information on the resulting adaptive cards, rendering a result similar to this:

![Screenshot of Copilot producing a bulleted list of products with references.](../media/3-06-b-find-northwind-beverages-with-more-than-100.png)

With this prompt, Copilot might also look in your OneDrive files to find the payment terms with each supplier's contract. In this case, you'll notice that some of the references won't have the **Northwind Inventory** icon, but the **Word** icon.

![Screenshot of Copilot extracting payment terms from contracts in SharePoint.](../media/3-06-c-payment-terms.png)

Here are some more prompts to try:

- _Find Northwind dairy products that are low on stock. Show me a table with the product, supplier, units in stock and on order._

- _We’ve been receiving partial orders for Tofu. Find the supplier in Northwind and draft an email summarizing our inventory and reminding them that they should stop sending partial orders per our MOQ policy._

- _Northwind will have a booth at Microsoft Community Days in London. Find products with local suppliers and write a LinkedIn post to promote the booth and products. Emphasize how delicious the products are and encourage people to attend our booth._

- _What beverage is high in demand due to social media that is low stock in Northwind in London. Reference the product details to update stock._

Which prompts work best for you? Try making up your own prompts and observe your log messages to see how Copilot accesses your plugin.

### Troubleshooting tip

If you're facing challenges while testing your plugin, you can enable **developer mode**. Developer mode provides information on the plugin selected by the Copilot orchestrator to respond to the prompt. It also shows the available functions in the plugin and the API call's status code.

To enable developer mode, type the following into Copilot:

```console
-developer on
```

Execute your prompt and developer mode outputs results that look similar to: 

![Screenshot of Copilot in developer mode.](../media/3-03-b-developer-mode.png)

As you can notice, below the response generated by Copilot, we have a table that provides us with insightful information about what happened behind the scenes:

- Under **Enabled plugins**, you can see that Copilot has identified that the Northwind Inventory plugin is enabled.

- Under **Matched functions**, you can see that Copilot has determined that the Northwind inventory plugin offers three functions: `inventorySearch`, `discountSearch`, and `companySearch`.

- Under **Selected functions for execution**, you can see that Copilot has selected the `inventorySearch` function to respond to the prompt.

- Under **Function execution details**, you can see some detailed information about the execution, like the HTTP response returned by the plugin to the Copilot engine.

## Check your work

After completing the tasks in this exercise, you should be able to use the **Northwind Inventory** plugin in Copilot for Microsoft 365. 

With that exercise complete, you're ready to add a new command to the messaging extension so that you can expand the plugin capabilities and perform more tasks. 

[Continue to the next exercise...](./5-exercise-3-add-new-command.md)
