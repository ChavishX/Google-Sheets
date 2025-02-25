function fetchWooCommerceData() {
  var sheetName = 'Woocommerce Waybill Data'; // Update this to match your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }
  
  // Calculate the start date for the last 7 days
  var now = new Date();
  var startDate = new Date(now);
  startDate.setDate(now.getDate() - 7);
  
  var apiUrl = 'https://nirasu.com/wp-json/wc/v3/orders';
  var consumerKey = 'ck_61ee4c25ce65b09fba16496ddba9eaeff6379ca0';
  var consumerSecret = 'cs_163bb9c6b14043fd3b40d34783681f41f380a650';
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Basic ' + Utilities.base64Encode(consumerKey + ':' + consumerSecret)
    }
  };
  
  // Add date filter to the API URL
  apiUrl += '?after=' + startDate.toISOString();
  
  var page = 1;
  var orders = [];
  var fetchedOrders;
  
  do {
    var paginatedUrl = apiUrl + '&page=' + page + '&per_page=100'; // Fetch 100 orders per page
    var response = UrlFetchApp.fetch(paginatedUrl, options);
    fetchedOrders = JSON.parse(response.getContentText());
    Logger.log('Fetched ' + fetchedOrders.length + ' orders from page ' + page);
    orders = orders.concat(fetchedOrders);
    page++;
  } while (fetchedOrders.length > 0);
  
  // Check if there are any orders fetched
  if (orders.length === 0) {
    Logger.log('No orders found for the last 7 days.');
    return;
  }
  
  // Get existing custom order numbers from the sheet
  var lastRow = sheet.getLastRow();
  var existingOrderNumbers = new Set();
  if (lastRow > 1) {
    existingOrderNumbers = new Set(sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat());
  }
  
  // Set headers if the sheet is empty
  if (lastRow === 0) {
    var headers = ['Order Date', 'Order Number', 'Customer Name', 'Billing Address', 'Default Address', 'Mobile Number', 'Metadata'];
    sheet.appendRow(headers);
  }
  
  // Sort orders by descending order date
  orders.sort(function(a, b) {
    return new Date(b.date_created) - new Date(a.date_created);
  });
  
  // Remove existing orders with the same order number
  orders.forEach(function(order) {
    var orderNumber = order.number;
    
    // Remove existing order if it already exists
    if (existingOrderNumbers.has(orderNumber)) {
      removeOrderByNumber(sheet, orderNumber);
    }
  });
  
  // Append order data
  orders.forEach(function(order) {
    var orderNumber = order.number;
    var metadata = order.meta_data.map(meta => meta.key + ': ' + meta.value).join(', ');
    
    var row = [
      order.date_created,
      orderNumber,
      order.billing.first_name + ' ' + order.billing.last_name,
      order.billing.address_1 + ', ' + order.billing.address_2 + ', ' + order.billing.city + ', ' + order.billing.state + ', ' + order.billing.postcode + ', ' + order.billing.country,
      order.shipping.address_1 + ', ' + order.shipping.address_2 + ', ' + order.shipping.city + ', ' + order.shipping.state + ', ' + order.shipping.postcode + ', ' + order.shipping.country,
      order.billing.phone,
      metadata
    ];
    sheet.appendRow(row);
    
    // Add the order number to the set
    existingOrderNumbers.add(orderNumber);
  });
}

function removeOrderByNumber(sheet, orderNumber) {
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) { // Start from the end to avoid index shifting
    if (data[i][1] === orderNumber) {
      sheet.deleteRow(i + 1); // Adjust for header row
    }
  }
}

function createTrigger() {
  // Create a time-driven trigger to run the fetchWooCommerceData function every 10 minutes
  ScriptApp.newTrigger('fetchWooCommerceData')
           .timeBased()
           .everyMinutes(10)
           .create();
}