---
lab:
    title: 'Exercise 2 - Update the API plugin definition'
    module: 'LAB 03: Use Adaptive Cards to show data in API plugins for declarative agents'
---

# Exercise 2 - Update the API plugin definition

The next step is to update the API plugin definition with Adaptive Cards that Copilot should use to display data from the API to users.

### Exercise Duration

- **Estimated Time to complete**: 10 minutes

## Task 1 - Add Adaptive Card to display a dish

In Visual Studio Code:

1. Open the **cards/dish.json** file and copy its contents.
1. Open the **appPackage/ai-plugin.json** file.
1. To the **functions.getDishes.capabilities.response_semantics** property, add a new property named **static_template** and set the **body** value to the contents of **dish.json**.
1. The complete code snippet looks like:

    ```json
    "static_template": {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.5",
      "body": [
          {
            "type": "Container",
            "items": [
              {
                "type": "Image",
                "url": "${image_url}",
                "size": "large"
              },
              {
                "type": "TextBlock",
                "text": "${name}",
                "weight": "Bolder"
              },
              {
                "type": "TextBlock",
                "text": "${description}",
                "wrap": true
              },
              {
                "type": "TextBlock",
                "text": "Allergens: ${if(count(allergens) > 0, join(allergens, ', '), 'none')}",
                "weight": "Lighter"
              },
              {
                "type": "TextBlock",
                "text": "**Price:** €${formatNumber(price, 2)}",
                "weight": "Lighter",
                "spacing": "None"
              }
            ]
          }
      ]
    }
    ```

1. Save your changes.

## Task 2 - Add Adaptive Card template to display the order summary

In Visual Studio Code:

1. Open the **cards/order.json** file and copy its contents.
1. Open the **appPackage/ai-plugin.json** file.
1. To the **functions.placeOrder.capabilities.response_semantics** property, add a new property named **static_template** and set its contents to the Adaptive Card.
1. The complete file looks like:

    ```json
    "static_template": {
      "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
      "type": "AdaptiveCard",
      "version": "1.5",
      "body": [
          {
            "type": "TextBlock",
            "text": "Order Confirmation 🤌",
            "size": "Large",
            "weight": "Bolder",
            "horizontalAlignment": "Center"
          },
          {
            "type": "Container",
            "items": [
              {
                "type": "TextBlock",
                "text": "Your order has been successfully placed!",
                "weight": "Bolder",
                "spacing": "Small"
              },
              {
                "type": "FactSet",
                "facts": [
                  {
                    "title": "Order ID:",
                    "value": "${order_id} "
                  },
                  {
                    "title": "Status:",
                    "value": "${status}"
                  },
                  {
                    "title": "Total Price:",
                    "value": "€${formatNumber(total_price, 2)}"
                  }
                ]
              }
            ]
          }
        ]
    }
    ```

1. Save your changes.
