# Zapier Alternative to Pull Orders from Shopify to Google Sheet

This repository contains a script that acts as an alternative to Zapier for pulling orders from a Shopify store into a Google Sheet. The script makes use of the Shopify API and Google Apps Script to automate the process.

## Prerequisites
Before you can use this script, ensure that you have completed the following steps:

1. Obtain your Shopify API credentials:
    - API Key: Replace `API_KEY` with your actual Shopify API key.
    - Password: Replace `PASSWORD` with your actual Shopify API password.
    - Store URL: Replace `SHOPIFY_STORE_URL` with the URL of your Shopify store.

2. Set the API version:
    - API Version: Replace `API_VERSION` with the desired Shopify API version.

3. Configure the starting order ID:
    - Starting Order ID: Replace `START_ORDER_ID` with the ID of the order from which you want to start retrieving orders. This step is optional, and if not provided, the script will start from the last retrieved order ID.

4. Set up a Google Sheet:
    - Create a Google Sheet in your Google Drive.
    - Create a sheet named 'ORDER RAW DATA' in the Google Sheet.

5. Create a trigger:
    - Run the `setStartingOrderId()` function manually once to set the starting order ID.
    - Create a time-based trigger for the `getShopifyOrders()` function to run every 30 minutes. This will automate the process of retrieving orders. Ensure to create the trigger using the Google Apps Script editor.

## Script Description
The script executes the following tasks:

1. Retrieves the last retrieved order ID from the script properties.
2. Constructs the Shopify API URL for retrieving orders.
3. Sends HTTP requests to the Shopify API, fetching orders in batches.
4. Concatenates the retrieved orders into an array.
5. Populates the orders into the 'ORDER RAW DATA' sheet in the Google Sheet.
6. Updates the last retrieved order ID in the script properties.
7. Repeats the process until all orders are retrieved.

## Important Note
To ensure smooth and automated workflow, it is essential to set up a trigger for the `setStartingOrderId()` function to run every 30 minutes. This trigger ensures that the script starts from the correct order ID when fetching orders.

For any further questions or concerns, please refer to the [GitHub repository](https://github.com/your-github-username/Zapier-Alternative-to-Pull-Orders-from-Shopify-to-Google-Sheet) (Replace with your repository link).

Feel free to customize the content and formatting of the `README.md` file according to your specific requirements.
