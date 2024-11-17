---
lab:
    title: 'Exercise 3 - Connect the plugin definition to the declarative agent'
    module: 'LAB 02: Build your first action for declarative agents with API plugin by using Visual Studio Code'
---

# Exercise 3 - Connect the plugin definition to the declarative agent

After you complete building the API plugin definition, the next step is to register it with the declarative agent. When users interact with the declarative agent, it matches the user's prompt against the defined API plugins and invokes the relevant functions.

### Exercise Duration

- **Estimated Time to complete**: 5 minutes

## Task 1 - Connect the plugin definition to the declarative agent

In Visual Studio Code:

1. Open the **appPackage/declarativeAgent.json** file.
1. After the **instructions** property, add the following code snippet:

    ```json
    "actions": [
      {
        "id": "menuPlugin",
        "file": "ai-plugin.json"
      }
    ]
    ```

    Using this snippet, you connect the declarative agent to the API plugin. You specify a unique ID for the plugin and instruct the agent where it can find the plugin's definition.

1. The complete file looks like:

    ```json
    {
      "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.0/schema.json",
      "version": "v1.0",
      "name": "Declarative agent",
      "description": "Declarative agent created with Teams Toolkit",
      "instructions": "$[file('instruction.txt')]",
      "actions": [
        {
          "id": "menuPlugin",
          "file": "ai-plugin.json"
        }
      ]
    }
    ```

1. Save your changes.

## Task 2 - Update declarative agent information and instruction

The declarative agent that you're building in this exercise helps users browse the menu of the local Italian restaurant and place orders. To optimize the agent for this scenario, update its name, description, and instructions.

In Visual Studio Code:

1. Update the declarative agent information:
    1. Open the **appPackage/declarativeAgent.json** file.
    1. Update the value of the **name** property to **Il Ristorante**.
    1. Update the value of the **description** property to **Order the most delicious Italian dishes and drinks from the comfort of your desk.**.
    1. Save the changes.
1. Update declarative agent's instructions:
    1. Open the **appPackage/instruction.txt** file.
    1. Replace its contents with:

        ```markdown
        You are an assistant specialized in helping users explore the menu of an Italian restaurant and place orders. You interact with the restaurant's menu API and guide users through the ordering process, ensuring a smooth and delightful experience. Follow the steps below to assist users in selecting their desired dishes and completing their orders:
        
        ### General Behavior:
        - Always greet the user warmly and offer assistance in exploring the menu or placing an order.
        - Use clear, concise language with a friendly tone that aligns with the atmosphere of a high-quality local Italian restaurant.
        - If the user is browsing the menu, offer suggestions based on the course they are interested in (breakfast, lunch, or dinner).
        - Ensure the conversation remains focused on helping the user find the information they need and completing the order.
        - Be proactive but never pushy. Offer suggestions and be informative, especially if the user seems uncertain.
        
        ### Menu Exploration:
        - When a user requests to see the menu, use the `GET /dishes` API to retrieve the list of available dishes, optionally filtered by course (breakfast, lunch, or dinner).
          - Example: If a user asks for breakfast options, use the `GET /dishes?course=breakfast` to return only breakfast dishes.
        - Present the dishes to the user with the following details:
          - Name of the dish
          - A tasty description of the dish
          - Price in € (Euro) formatted as a decimal number with two decimal places
          - Allergen information (if relevant)
          - Don't include the URL.
        
        ### Beverage Suggestion:
        - If the order does not already include a beverage, suggest a suitable beverage option based on the course.
        - Use the `GET /dishes?course={course}&type=drink` API to retrieve available drinks for that course.
        - Politely offer the suggestion: *"Would you like to add a beverage to your order? I recommend [beverage] for [course]."*
        
        ### Placing the Order:
        - Once the user has finalized their order, use the `POST /order` API to submit the order.
          - Ensure the request includes the correct dish names and quantities as per the user's selection.
          - Example API payload:
         
            ```json
            {
              "dishes": [
                {
                  "name": "frittata",
                  "quantity": 2
                },
                {
                  "name": "cappuccino",
                  "quantity": 1
                }
              ]
            }
            ```
        
        ### Error Handling:
        - If the user selects a dish that is unavailable or provides an invalid dish name, respond gracefully and suggest alternative options.
          - Example: *"It seems that dish is currently unavailable. How about trying [alternative dish]?"*
        - Ensure that any errors from the API are communicated politely to the user, offering to retry or explore other options.
        ```

        Notice that in the instructions, we define the general behavior of the agent and instruct it what it's capable of. We also include instructions for specific behavior around placing an order, including the shape of the data that the API expects. We include this information to ensure that the agent works as intended.

    > [!NOTE]
    > To preserve formatting, you may need to perform multiple copy/paste operations into Notepad before copying into Visual Studio Code.

    1. Save the changes.
1. To help users, understand what they can use the agent for, add conversation starters:
    1. Open the **appPackage/declarativeAgent.json** file.
    1. After the **instructions** property, add a new property named **conversation_starters**:

        ```json
        "conversation_starters": [
          {
            "text": "What's for lunch today?"
          },
          {
            "text": "What can I order for dinner that is gluten-free?"
          }
        ]
        ```

    1. The complete file looks like:

        ```json
        {
          "$schema": "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.0/schema.json",
          "version": "v1.0",
          "name": "Il Ristorante",
          "description": "Order the most delicious Italian dishes and drinks from the comfort of your desk.",
          "instructions": "$[file('instruction.txt')]",
          "conversation_starters": [
            {
              "text": "What's for lunch today?"
            },
            {
              "text": "What can I order for dinner that is gluten-free?"
            }
          ],
          "actions": [
            {
              "id": "menuPlugin",
              "file": "ai-plugin.json"
            }
          ]
        }
        ```

    1. Save your changes.

## Task 3 - Update the API URL

Before you can test your declarative agent, you need to update the URL of the API in the API specification file. Right now, the URL is set to `http://localhost:7071/api` which is the URL that Azure Functions uses when running locally. However, because you want Copilot to call your API from the cloud, you need to expose your API to the internet. Teams Toolkit automatically exposes your local API over the internet, by creating a dev tunnel. Each time you start debugging your project, Teams Toolkit starts a new dev tunnel and stores its URL in the **OPENAPI_SERVER_URL** variable. You can see how Teams Toolkit starts the tunnel and stores its URL in the **.vscode/tasks.json** file, in the **Start local tunnel** task:

```json
{
  // Start the local tunnel service to forward public URL to local port and inspect traffic.
  // See https://aka.ms/teamsfx-tasks/local-tunnel for the detailed args definitions.
  "label": "Start local tunnel",
  "type": "teamsfx",
  "command": "debug-start-local-tunnel",
  "args": {
    "type": "dev-tunnel",
    "ports": [
      {
        "portNumber": 7071,
        "protocol": "http",
        "access": "public",
        "writeToEnvironmentFile": {
          "endpoint": "OPENAPI_SERVER_URL", // output tunnel endpoint as OPENAPI_SERVER_URL
        }
      }
    ],
    "env": "local"
  },
  "isBackground": true,
  "problemMatcher": "$teamsfx-local-tunnel-watch"
}
```

To use this tunnel, you need to update your API specification to use the **OPENAPI_SERVER_URL** variable.

In Visual Studio Code:

1. Open the **appPackage/apiSpecificationFile/ristorante.yml** file.
1. Change the value of the **servers.url** property to **${{OPENAPI_SERVER_URL}}/api**.
1. The changed file looks like:

    ```yaml
    openapi: 3.0.0
    info:
      title: Il Ristorante menu API
      version: 1.0.0
      description: API to retrieve dishes and place orders for Il Ristorante.
    servers:
      - url: ${{OPENAPI_SERVER_URL}}/api
        description: Il Ristorante API server
    paths:
      ...trimmed for brevity
    ```

1. Save your changes.

Your API plugin is finished and integrated with a declarative agent. Continue with testing the agent in Microsoft 365 Copilot.