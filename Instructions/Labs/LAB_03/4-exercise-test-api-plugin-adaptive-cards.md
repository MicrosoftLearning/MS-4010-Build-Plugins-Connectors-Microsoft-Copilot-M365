---
lab:
    title: 'Exercise 3 - Test the declarative agent with the API plugin in Microsoft 365 Copilot'
    module: 'LAB 03: Use Adaptive Cards to show data in API plugins for declarative agents'
---

# Exercise 3 - Test the declarative agent with the API plugin in Microsoft 365 Copilot

The final step is to test the declarative agent with API plugin in Microsoft 365 Copilot.

### Exercise Duration

- **Estimated Time to complete**: 10 minutes

## Task 1 - Provision and start debugging debug

In Visual Studio Code:

1. From the **Activity Bar**, choose **Teams Toolkit**.
1. In the **Accounts** section, ensure that you're signed in to your Microsoft 365 tenant with Microsoft 365 Copilot.

    ![Screenshot of the Teams Toolkit accounts section in Visual Studio Code.](../media/LAB_03/3-teams-toolkit-accounts.png)

1. From the **Activity Bar**, choose **Run and Debug**.
1. Select the **Debug in Copilot** configuration and start debugging using the **Start Debugging** button.  

    ![Screenshot of the Debug in Copilot configuration in Visual Studio Code.](../media/LAB_03/3-visual-studio-code-start-debugging.png)

1. Visual Studio Code builds and deploys your project to your Microsoft 365 tenant and opens a new web browser window.

## Task 2 - Test and review results

In the web browser:

1. When prompted, sign in with the account that belongs to your Microsoft 365 tenant with Microsoft 365 Copilot.
1. From the side bar, select **Il Ristorante**.

    ![Screenshot of the Microsoft 365 Copilot interface with the Il Ristorante agent selected.](../media/LAB_03/3-copilot-select-agent.png)

1. Choose the **What's for lunch today?** conversation starter and submit the prompt.

    ![Screenshot of the Microsoft 365 Copilot interface with the lunch prompt.](../media/LAB_03/3-copilot-lunch-prompt.png)

1. When prompted, examine the data that the agent sends to the API and confirm using the **Allow once** button.

    ![Screenshot of the Microsoft 365 Copilot interface with the lunch confirmation.](../media/LAB_03/3-copilot-lunch-confirm.png)

1. Wait for the agent to respond. Notice that the popup on a citation now includes your custom Adaptive Card with additional information from the API.

    ![Screenshot of the Microsoft 365 Copilot interface with the lunch response.](../media/LAB_03/3-copilot-lunch-response.png)

1. Place an order, by typing in the prompt text box: **1x spaghetti, 1x iced tea** and submit the prompt.
1. Examine the data that the agent sends to the API and continue using the **Confirm** button.

    ![Screenshot of the Microsoft 365 Copilot interface with the order confirmation.](../media/LAB_03/3-copilot-order-confirm.png)

1. Wait for the agent to place the order and return the order summary. Notice, that because the API returns a single item, the agent renders it using an Adaptive Card and includes the card directly in its response.

    ![Screenshot of the Microsoft 365 Copilot interface with the order response.](../media/LAB_03/3-copilot-order-response.png)

1. Go back to Visual Studio Code and stop debugging.
1. Switch to the **Terminal** tab and close all active terminals.

    ![Screenshot of the Visual Studio Code terminal tab with the option to close all terminals.](../media/LAB_03/3-visual-studio-code-close-terminal.png)
