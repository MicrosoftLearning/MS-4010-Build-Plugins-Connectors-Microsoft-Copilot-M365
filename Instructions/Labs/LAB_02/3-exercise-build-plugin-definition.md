---
lab:
    title: 'Exercise 3 - Build API plugin definition'
    module: 'LAB 02: Build your first action for declarative agents with API plugin by using Visual Studio Code'
---

# Exercise 2 - Build API plugin definition

The next step is to add the plugin definition to the project. The plugin definition contains the following information:

- What actions the plugin can perform.
- What's the shape of the data that it expects and returns.
- How the declarative agent must call the underlying API.

### Exercise Duration

- **Estimated Time to complete**: 10 minutes

## Task 1 - Add the basic plugin definition structure

In Visual Studio Code:

1. In the **appPackage** folder, add a new file named **ai-plugin.json**.
1. Paste the following contents:

    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.1/schema.json",
      "schema_version": "v2.1",
      "namespace": "ilristorante",
      "name_for_human": "Il Ristorante",
      "description_for_human": "See the today's menu and place orders",
      "description_for_model": "Plugin for getting the today's menu, optionally filtered by course and allergens, and placing orders",
      "functions": [ 
      ],
      "runtimes": [
      ],
      "capabilities": {
        "localization": {},
        "conversation_starters": []
      }
    }
    ```

    The file contains a basic structure for an API plugin with a description for the human and the model. The **description_for_model** includes detailed information about what the plugin can do to help the agent understand when it should consider invoking it.
1. Save your changes.

## Task 2 - Define functions

An API plugin defines one or more functions that map to API operations defined in the API specification. Each function consists of a name and a description and a response definition that instructs the agent how to display the data to users.

### Define a function to retrieve the menu

Start with defining a function to retrieve the information about today's menu.

In Visual Studio Code:

1. Open the **appPackage/ai-plugin.json** file.
1. In the **functions** array, add the following snippet:

    ```json
    {
      "name": "getDishes",
      "description": "Returns information about the dishes on the menu. Can filter by course (breakfast, lunch or dinner), name, allergens, or type (dish, drink).",
      "capabilities": {
        "response_semantics": {
          "data_path": "$.dishes",
          "properties": {
            "title": "$.name",
            "subtitle": "$.description"
          }
        }
      }
    }
    ```

    You start by defining a function that invokes the **getDishes** operation from the API specification. Next, you provide a function description. This description is important because Copilot uses it to decide which function to invoke for a user's prompt.

    In the response_semantics property, you specify how Copilot should display the data it receives from the API. Because the API returns the information about the dishes on the menu in the **dishes** property, you set the **data_path** property to the `$.dishes` JSONPath expression.

    Next, in the **properties** section, you map which properties from the API response represent the title, description, and URL. Because in this case the dishes don't have a URL, you only map the **title** and **description**.

1. The complete code snippet looks like:

    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.1/schema.json",
      "schema_version": "v2.1",
      "namespace": "ilristorante",
      "name_for_human": "Il Ristorante",
      "description_for_human": "See the today's menu and place orders",
      "description_for_model": "Plugin for getting the today's menu, optionally filtered by course and allergens, and placing orders",
      "functions": [
        {
          "name": "getDishes",
          "description": "Returns information about the dishes on the menu. Can filter by course (breakfast, lunch or dinner), name, allergens, or type (dish, drink).",
          "capabilities": {
            "response_semantics": {
              "data_path": "$.dishes",
              "properties": {
                "title": "$.name",
                "subtitle": "$.description"
              }
            }
          }
        }
      ],
      "runtimes": [
      ],
      "capabilities": {
        "localization": {},
        "conversation_starters": []
      }
    }
    ```

1. Save your changes.

### Define a function to place the order

Next, define a function to place the order.

In Visual Studio Code:

1. Open the **appPackage/ai-plugin.json** file.
1. To the end of the **functions** array, add the following snippet:

    ```json
    {
      "name": "placeOrder",
      "description": "Places an order and returns the order details",
      "capabilities": {
        "response_semantics": {
          "data_path": "$",
          "properties": {
            "title": "$.order_id",
            "subtitle": "$.total_price"
          }
        }
      }
    }
    ```

    You start by referring to the API operation with ID **placeOrder**. Then, you provide a description that Copilot uses to match this function against a user's prompt. Next, you instruct Copilot how to return the data. Following is the data that the API returns after placing an order:

    ```json
    {
      "order_id": 6532,
      "status": "confirmed",
      "total_price": 21.97
    }
    ```

    Because the data that you want to show is located directly in the root of the response object, you set the **data_path** to **$** which indicates the top node of the JSON object. You define the title to show the order's number and the subtitle its price.

1. The complete file looks like:

    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.1/schema.json",
      "schema_version": "v2.1",
      "namespace": "ilristorante",
      "name_for_human": "Il Ristorante",
      "description_for_human": "See the today's menu and place orders",
      "description_for_model": "Plugin for getting the today's menu, optionally filtered by course and allergens, and placing orders",
      "functions": [
        {
          "name": "getDishes",
          "description": "Returns information about the dishes on the menu. Can filter by course (breakfast, lunch or dinner), name, allergens, or type (dish, drink).",
          ...trimmed for brevity
        },
        {
          "name": "placeOrder",
          "description": "Places an order and returns the order details",
          "capabilities": {
            "response_semantics": {
              "data_path": "$",
              "properties": {
                "title": "$.order_id",
                "subtitle": "$.total_price"
              }
            }
          }
        }
      ],
      "runtimes": [
      ],
      "capabilities": {
        "localization": {},
        "conversation_starters": []
      }
    }
    ```

1. Save your changes.

## Task 3 - Define runtimes

After defining functions for Copilot to invoke, the next step is to instruct it how it should call them. You do that in the **runtimes** section of the plugin definition.

In Visual Studio Code:

1. Open the **appPackage/ai-plugin.json** file.
1. In the runtimes array, add the following code:

    ```json
    {
      "type": "OpenApi",
      "auth": {
        "type": "None"
      },
      "spec": {
        "url": "apiSpecificationFile/ristorante.yml"
      },
      "run_for_functions": [
        "getDishes",
        "placeOrder"
      ]
    }
    ```

    You start with instructing Copilot that you provide it with OpenAPI information about the API (**type: OpenApi**) to call and that it's anonymous (**auth.type: None**). Next, in the **spec** section, you specify the relative path to the API specification located in your project. Finally, in the **run_for_functions** property, you list all functions that belong to this API.

1. The complete file looks like:

    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.1/schema.json",
      "schema_version": "v2.1",
      "namespace": "ilristorante",
      "name_for_human": "Il Ristorante",
      "description_for_human": "See the today's menu and place orders",
      "description_for_model": "Plugin for getting the today's menu, optionally filtered by course and allergens, and placing orders",
      "functions": [
        {
          "name": "getDishes",
          ...trimmed for brevity
        },
        {
          "name": "placeOrder",
          ...trimmed for brevity
        }
      ],
      "runtimes": [
        {
          "type": "OpenApi",
          "auth": {
            "type": "None"
          },
          "spec": {
            "url": "apiSpecificationFile/ristorante.yml"
          },
          "run_for_functions": [
            "getDishes",
            "placeOrder"
          ]
        }
      ],
      "capabilities": {
        "localization": {},
        "conversation_starters": []
      }
    }
    ```

1. Save your changes.

