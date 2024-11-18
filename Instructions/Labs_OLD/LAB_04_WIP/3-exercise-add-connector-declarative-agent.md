---
lab:
    title: 'Exercise 2 - Create declarative agent and integrate Graph connector'
    module: 'LAB 04: Add custom knowledge to declarative agents using Microsoft Graph connectors and Visual Studio Code'
---

# Exercise 2 - Create declarative agent and integrate Graph connector


### Exercise Duration

- **Estimated Time to complete**: 10 minutes

## Task 1 - Create a new declarative agent

One way to create a declarative agent is by using Teams Toolkit. Teams Toolkit offers a template project for creating declarative agents, which gives you a great place to start with configuring your agent's settings and including extra capabilities.

To create a new declarative agent, open Visual Studio Code.

In Visual Studio Code:

1. In the **Activity Bar** (side bar), open the **Teams Toolkit** extension.
1. In the Teams Toolkit pane, select the **Create a New App** button.
1. In the **New Project** dialog, choose the **Copilot Agent** option.
1. In the next dialog, choose the **Declarative Agent** option.
1. Don't add a plugin by selecting the **No plugin** option.
1. Select a folder where you want to store the project on your computer.
1. Name your project `da-it-policies`.

## Task 2 - Configure declarative agent instructions

Teams Toolkit creates a new declarative agent project. To scope it to your scenario, update the agent's description and instructions.

In Visual Studio Code:

1. Open the **appPackage/declarativeAgent.json** file
1. Update the value of the **name** property to `Contoso IT policies`
1. Update the value of the **description** property to `Assistant specialized in Contoso IT policies`.
1. Save your changes.
1. The updated file contents look as follows:

    ```json
    {
      "$schema": "https://aka.ms/json-schemas/copilot/declarative-agent/v1.0/schema.json",
      "name": "Contoso IT policies",
      "description": "Assistant specialized in Contoso IT policies",
      "instructions": "$[file('instruction.txt')]",
    }
    ```

1. Open the **appPackage/instruction.txt** file.
1. Update the contents to:

    ```text
    You are a helpful assistant that can answer questions from users in a friendly manner. You should take your time to respond. Your responses should be accurate and helpful. If you don't have the information, you should say so in your response. When answering follow-up questions, you should review the information you gathered from external sources, if you don't already have the information to give an accurate answer, you should search for more information. Only answer when you have the information to give an accurate response.
    ```

1. Save your changes.

## Task 3 - Integrate Microsoft Graph connector with a declarative agent

After you create a declarative agent, the next step is to integrate it with a Microsoft Graph connector so that it can access external data.

In Visual Studio Code:

1. Open the **appPackage/declarativeAgent.json** file.
1. After the **instructions** property, add a new property named **capabilities** with the following code:

    ```json
    { 
      "$schema": "https://aka.ms/json-schemas/copilot/declarative-agent/v1.0/schema.json",
      "name": "Contoso IT policies",
      "description": "Assistant specialized in Contoso IT policies",
      "instructions": "$[file('instruction.txt')]",
      "capabilities": [
        {
          "name": "GraphConnectors",
          "connections": [ 
            {
              "connection_id": ""
            }
          ]
        }
      ]
    } 
    ```

1. In the **connection_id** property, specify `policieslocal` as the ID of the external connection. `policieslocal` is the ID of the external connection that the Graph connector created in the previous steps.
1. The updated file contents look as follows:

    ```json
    { 
      "$schema": "https://aka.ms/json-schemas/copilot/declarative-agent/v1.0/schema.json",
      "name": "Contoso IT policies",
      "description": "Assistant specialized in Contoso IT policies",
      "instructions": "$[file('instruction.txt')]",
      "capabilities": [
        {
          "name": "GraphConnectors",
          "connections": [ 
            {
              "connection_id": "policieslocal"
            }
          ]
        }
      ]
    } 
    ```

1. Save your changes.
