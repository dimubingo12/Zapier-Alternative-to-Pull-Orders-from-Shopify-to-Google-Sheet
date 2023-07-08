Shopify Order Retrieval Script
This script is designed to retrieve orders from a Shopify store and populate them into a Google Sheet. It utilizes the Shopify API and Google Apps Script to automate the process. The script is intended to be used as an alternative to Zapier for pulling orders from Shopify into a Google Sheet.

Prerequisites
Before using this script, please ensure you have completed the following steps:

Obtain your Shopify API credentials:

API Key: Replace API_KEY with your actual Shopify API key.
Password: Replace PASSWORD with your actual Shopify API password.
Store URL: Replace SHOPIFY_STORE_URL with the URL of your Shopify store.
Set the API version:

API Version: Replace API_VERSION with the desired Shopify API version.
Configure the starting order ID:

Starting Order ID: Replace START_ORDER_ID with the ID of the order from which you want to start retrieving orders. This is optional, and if not provided, the script will start from the last retrieved order ID.
Set up a Google Sheet:

Create a Google Sheet in your Google Drive.
Create a sheet named 'ORDER RAW DATA' in the Google Sheet.
Create a trigger:

Run the setStartingOrderId() function manually once to set the starting order ID.
Create a time-based trigger for the getShopifyOrders() function to run every 30 minutes. This will automate the process of retrieving orders. Make sure to create the trigger using the Google Apps Script editor.
Script Description
The script performs the following tasks:

Retrieves the last retrieved order ID from the script properties.
Constructs the Shopify API URL for retrieving orders.
Sends HTTP requests to the Shopify API, fetching orders in batches.
Concatenates the retrieved orders into an array.
Populates the orders into the 'ORDER RAW DATA' sheet in the Google Sheet.
Updates the last retrieved order ID in the script properties.
Repeats the process until all orders are retrieved.

Important Note
To ensure a smooth and automated flow, it is essential to set up a trigger for the setStartingOrderId() function to run every 30 minutes. This trigger ensures that the script starts from the correct order ID when fetching orders.


For any further questions or concerns, please refer to the GitHub repository Zapier-Alternative-to-Pull-Orders-from-Shopify-to-Google-Sheet.
