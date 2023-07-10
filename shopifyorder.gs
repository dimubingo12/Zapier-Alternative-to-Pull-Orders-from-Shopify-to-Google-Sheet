var API_KEY = "bxxxxxxxxxxxxxx";
var PASSWORD = "shpat_xxxxxxxxxxxxx";
var SHOPIFY_STORE_URL = "https://xxxxxx.myshopify.com";
var API_VERSION = "2023-04";
var LAST_ORDER_ID_PROPERTY = 'lastOrderIdlll';
var START_ORDER_ID = "5xxxxxx"; 
var LAST_POPULATED_ROW_PROPERTY = 'lastPopulatedRow';

function getShopifyOrders() {
  var lastOrderId = PropertiesService.getScriptProperties().getProperty(LAST_ORDER_ID_PROPERTY);
  var lastPopulatedRow = PropertiesService.getScriptProperties().getProperty(LAST_POPULATED_ROW_PROPERTY);
  if (lastOrderId == null) {
    lastOrderId = START_ORDER_ID;
    PropertiesService.getScriptProperties().setProperty(LAST_ORDER_ID_PROPERTY, lastOrderId);
  }
  var orders = [];
  var url = SHOPIFY_STORE_URL + "/admin/api/" + API_VERSION + "/orders.json?status=any";
  if (lastOrderId != null) {
    url += "&since_id=" + lastOrderId;
  }

  var options = {
    "headers": {
      "Authorization": "Basic " + Utilities.base64Encode(API_KEY + ":" + PASSWORD)
    }
  };

  do {
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    orders = orders.concat(data.orders);
    url = getNextPageUrl(response.getHeaders()['Link']);
  } while (url != null);

  if (orders.length > 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ORDER RAW DATA');
    var startRow = lastPopulatedRow ? parseInt(lastPopulatedRow) + 1 : 2094;
    var rowsToUpdate = [];
    var lastPopulatedRow = startRow - 1;
    


orders.forEach(function (order) {
      var row = new Array(50).fill('');
 
  var billingName = order.billing_address ? order.billing_address.first_name + " " + order.billing_address.last_name : 'Not Provided';
  var row = new Array(50).fill('');  // initialize an array of 45 elements all set to an empty string
  row[0] = order.created_at; // Column A - Created At
  row[2] = order.name;       // Column C - Name
  row[3] = billingName;      // Column D - Billing First Name and Last Name
  row[4] = order.financial_status; // Column E - Financial status
  row[5] = order.closed_at; // Column E - Financial status
  row[8] = order.buyer_accepts_marketing; // Column F - customer accepted marketing
  row[9] = order.currency; // Column G - fulfillment status
  row[10] = order.current_subtotal_price; // Column G - fulfillment status
  row[11] = order.shipping_lines[0].price; 
  row[12] = order.current_total_price;
  row[13] = order.discount_codes && order.discount_codes.length > 0 ? order.discount_codes[0].code : ''; // discount_codes code
  row[14] = order.discount_codes && order.discount_codes.length > 0 ? order.discount_codes[0].amount : 0; // discount_codes amount
  row[15] = order.shipping_lines && order.shipping_lines.length > 0 ? order.shipping_lines[0].code : ''; // shipping_lines code
  row[16] = order.line_items ? order.line_items.map(item => item.quantity.toString()).join(', ') : ''; // Concatenates all line items quantities separated by commas
  row[17] = order.line_items ? order.line_items.map(item => item.name).join(', ') : ''; // Concatenates all line items names separated by commas
  row[18] = order.line_items ? order.line_items.map(item => item.price).join(', ') : ''; // Concatenates all line items prices separated by commas
  var lineItemPrices = order.line_items.map(function(lineItem) {
  return lineItem.price_set.shop_money.amount;
}).join(", ");

  row[19] = lineItemPrices; // Column 21 - Line Items Price (Set Shop Money Amount)

  row[20] = order.line_items ? order.line_items.map(item => item.sku).join(', ') : ''; // Concatenates all line items SKUs separated by commas
  row[21] = order.billing_address ? order.billing_address.address1 : ''; // Column 21 - Billing Address 1
  row[22] = order.billing_address ? order.billing_address.address1 : ''; // Column 22 - Billing Address 1
  row[23] = order.billing_address ? order.billing_address.address2 : ''; // Column 23 - Billing Address 1
  row[24] = order.billing_address ? order.billing_address.city : ''; // Column 24 - Billing Address 1
  row[25] = order.billing_address ? order.billing_address.zip : ''; // Column 25 - Billing Address 1
  row[26] = order.billing_address ? order.billing_address.province : ''; // Column 26 - Billing Address 1
  row[27] = order.billing_address ? order.billing_address.country_code : ''; // Column 27 - Billing Address 1
  row[28] = order.billing_address ? order.billing_address.phone : ''; // Column 28 - Billing Address 1
  row[29] = order.billing_address ? order.shipping_address.phone : ''; // Column 29 - Billing Address 1
  var noteAttrs = order.note_attributes ? order.note_attributes.map(function(attr) {
  return attr.name + ": " + attr.value;
}).join(", ") : '';

  row[31] = noteAttrs; // Column 30 - Note Attributes
var paymentGateways = order.payment_gateway_names ? order.payment_gateway_names.join(", ") : '';

  row[32] = paymentGateways; // Column 32 - Payment Gateway Names
 var paymentReference = order.token ? order.token : '';

  row[33] = paymentReference; // Column 34 - Payment Reference

var email = order.email ? order.email : '';

  row[34] = email; // Column 34 - Email
var userAgent = order.client_details ? order.client_details.user_agent : '';
  row[44] = order.id;        // Column AS - Order ID
  row[45] = userAgent; // Column 35 - User Agent
  var landingSite = order.landing_site ? order.landing_site : '';
  row[46] = landingSite; // Column 36 - Landing Site


    rowsToUpdate.push(row);
     lastPopulatedRow++;
    });

    var numRows = rowsToUpdate.length;
    var numColumns = rowsToUpdate[0].length;
    var range = sheet.getRange(startRow, 1, numRows, numColumns);
    var existingData = range.getValues();

    for (var i = 0; i < numRows; i++) {
      var rowData = rowsToUpdate[i];
      var existingRowData = existingData[i];

      for (var j = 0; j < numColumns; j++) {
        var cell = sheet.getRange(startRow + i, j + 1);
        var formula = cell.getFormula();
        var isFormulaArray = formula && formula.indexOf("ARRAYFORMULA") !== -1;

        // Skip columns AY (51) and AZ (52)
        if (j === 51 || j === 52) {
          continue;
        }

        if (existingRowData[j] === '' && !isFormulaArray) {
          cell.setValue(rowData[j]);
        }
      }
    }

    var currentOrderId = orders[orders.length - 1].id.toString();
    PropertiesService.getScriptProperties().setProperty(LAST_ORDER_ID_PROPERTY, currentOrderId);
    PropertiesService.getScriptProperties().setProperty(LAST_POPULATED_ROW_PROPERTY, lastPopulatedRow.toString());
  }
}




function getNextPageUrl(linkHeader) {
  if (!linkHeader) {
    return null;
  }

  var links = linkHeader.split(',');
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    if (link.indexOf('rel="next"') > -1) {
      return link.slice(link.indexOf("<") + 1, link.indexOf(">"));
    }
  }

  return null;
}


function createTriggersp() {
  ScriptApp.newTrigger('getShopifyOrders')
    .timeBased()
    .everyMinutes(30)
    .create();
}

// Call the getShopifyOrders function initially to populate the data
getShopifyOrders();


