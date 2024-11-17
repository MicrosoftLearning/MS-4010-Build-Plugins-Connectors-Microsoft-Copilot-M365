---
lab:
    title: 'Exercise 1 - Download project and examine files'
    module: 'LAB 02: Build your first action for declarative agents with API plugin by using Visual Studio Code'
---

# Exercise 1 - Download project and examine files

Extending a declarative agent with actions allows it to retrieve and update data stored in external systems in real-time. Using API plugins, you can connect to external systems through their APIs to retrieve and update information.

### Exercise Duration

- **Estimated Time to complete**: 10 minutes

## Task 1 - Download the starter project

Start by downloading the sample project. In a web browser:

1. Navigate to [https://github.com/microsoft/learn-declarative-agent-api-plugin-typescript](https://github.com/microsoft/learn-declarative-agent-api-plugin-typescript).
    1. Follow the steps to [download the repository source code](https://docs.github.com/repositories/working-with-files/using-files/downloading-source-code-archives#downloading-source-code-archives-from-the-repository-view) to your computer.
    1. Extract the contents of the downloaded ZIP file and extract it to your **Documents folder**.
    1. Open the folder in Visual Studio Code.

The sample project is a Teams Toolkit project that includes a declarative agent and an anonymous API running on Azure Functions. The declarative agent is identical to a newly created declarative agent using Teams Toolkit. The API belongs to a fictitious Italian restaurant and allows you to browse today's menu and place orders.

## Task 2 - Examine the API definition

First, look at the API definition of the Italian restaurant's API.

In Visual Studio Code:

1. In the **Explorer** view, open the **appPackage/apiSpecificationFile/ristorante.yml** file. The file is an OpenAPI specification that describes the API of the Italian restaurant.
1. Locate the **servers.url** property

    ```yaml
    servers:
      - url: http://localhost:7071/api
        description: Il Ristorante API server
    ```

    Notice that it's pointing to a local URL which matches the standard URL when running Azure Functions locally.

1. Locate the **paths** property, which contains two operations: **/dishes** for retrieving today's menu, and **/orders** for placing an order.

    > [!IMPORTANT]
    > Notice that each operation contains the **operationId** property that uniquely identifies the operation in the API specification. Copilot requires each operation to have a unique ID so that it knows which API it should call for specific user prompts.

## Task 3 - Examine the API implementation

Next, look at the sample API that you use in this exercise.

In Visual Studio Code:

1. In the **Explorer** view, open the **src/data.json** file. The file contains fictitious menu item for our Italian restaurant. Each dish consists of:

    - name,
    - description,
    - link to an image,
    - price,
    - in which course it's served,
    - type (dish or drink),
    - optionally a list of allergens

    In this exercise, APIs use this file as their data source.
1. Next, expand the **src/functions** folder. Notice two files named **dishes.ts** and **placeOrder.ts**. These files contain implementation of the two operations defined in the API specification.
1. Open the **src/functions/dishes.ts** file. Take a moment to review how the API is working. It starts with loading the sample data from the **src/functions/data.json** file.

    ```typescript
    import data from "../data.json";
    ```

    Next, it looks in the different query string parameters for possible filters that the client calling the API might pass.

    ```typescript
    const course = req.query.get('course');
    const allergensString = req.query.get('allergens');
    const allergens: string[] = allergensString ? allergensString.split(",") : [];
    const type = req.query.get('type');
    const name = req.query.get('name');
    ```

    Based on the filters specified on the request, the API filters the data set and returns a response.

1. Next, examine the API for placing orders defined in the **src/functions/placeOrder.ts** file. The API starts with referencing the sample data. Then, it defines the shape of the order that the client sends in the request body.

    ```typescript
    interface OrderedDish {
      name?: string;
      quantity?: number;
    }
    
    interface Order {
      dishes: OrderedDish[];
    }
    ```

    When the API processes the request, it first checks if the request contains a body and if it has the correct shape. If not, it rejects the request with a 400 Bad Request error.

    ```typescript
    let order: Order | undefined;
    try {
      order = await req.json() as Order | undefined;
    }
    catch (error) {
      return {
        status: 400,
        jsonBody: { message: "Invalid JSON format" },
      } as HttpResponseInit;
    }
    
    if (!order.dishes || !Array.isArray(order.dishes)) {
      return {
        status: 400,
        jsonBody: { message: "Invalid order format" }
      } as HttpResponseInit;
    }
    ```

    Next, the API resolves the request into dishes on the menu and calculates the total price.

    ```typescript
    let totalPrice = 0;
    const orderDetails = order.dishes.map(orderedDish => {
      const dish = data.find(d => d.name.toLowerCase().includes(orderedDish.name.toLowerCase()));
      if (dish) {
        totalPrice += dish.price * orderedDish.quantity;
        return {
          name: dish.name,
          quantity: orderedDish.quantity,
          price: dish.price,
        };
      }
      else {
        context.error(`Invalid dish: ${orderedDish.name}`);
        return null;
      }
    });
    ```

    > [!IMPORTANT]
    > Notice how the API expects the client to specify the dish by a part of its name rather than its ID. This is on purpose because large language models work better with words than numbers. Additionally, before calling the API to place the order, Copilot has the name of the dish readily available as part of the user's prompt. If Copilot had to refer to a dish by its ID, it would first need to retrieve which requires additional API requests and which Copilot can't do now.

    When the API is ready, it returns a response with a total price, and a made-up order ID, and status.

    ```typescript
    const orderId = Math.floor(Math.random() * 10000);
    
    return {
      status: 201,
      jsonBody: {
        order_id: orderId,
        status: "confirmed",
        total_price: totalPrice,
      }
    } as HttpResponseInit;
    ```

