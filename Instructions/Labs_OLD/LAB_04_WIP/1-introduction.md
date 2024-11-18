---
lab:
    title: 'Introduction'
    module: 'LAB 04: Add custom knowledge to declarative agents using Microsoft Graph connectors and Visual Studio Code'
---

# Introduction

Microsoft 365 Copilot agents let you create AI-powered assistants optimized for specific scenarios. Using instructions, you define the context for the copilot and configure its tone of voice or how it should respond. By configuring the agent's knowledge, you give it access to external data that isn't a part of the Large Language Model (LLM), so that it can respond more accurately. 

## Example scenario

Suppose you work in an IT department in a large organization. Your organization standardizes IT through different policies which it stores in a specialized system. You and your colleagues in the IT department regularly get questions that are covered in policies. Looking up answers in the policies management system is time-consuming. You would like to provide your organization with an AI-powered assistant capable of answering your colleagues' questions using authoritative information from the policies.

## Learning objectives

By the end of this module, you'll be able to build declarative agents for Microsoft 365 Copilot. You'll understand how to configure their instructions to optimize them for a specific scenario. You'll also know how to integrate them with Microsoft Graph connectors to give them access to external data, that's not a part of the Microsoft 365 Copilot's LLM.

## Prerequisites

- Knowledge of what Microsoft 365 Copilot is and how it works at the beginner level
- Knowledge of how to build a Microsoft 365 Copilot declarative agent
- Knowledge of how to build a Graph connector
- Microsoft 365 tenant with Microsoft 365 Copilot, and tenant administrator privileges
- [Visual Studio Code](https://code.visualstudio.com/) with the [Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) extension installed
- [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local#install-the-azure-functions-core-tools)
- [Node.js v18](https://nodejs.org/)
