function fetchWooCommerceData() {
  var sheetName = 'Woocommerce Monthly Data'; // Update this to match your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }
  
  // Clear the sheet before fetching new data
  sheet.clearContents();
  
  // Calculate the start of the current month
  var now = new Date();
  var startDate = new Date(now.getFullYear(), now.getMonth(), 1);
  
  var apiUrl = 'https://nirasu.com/wp-json/wc/v3/orders';
  var consumerKey = 'ck_61ee4c25ce65b09fba16496ddba9eaeff6379ca0';
  var consumerSecret = 'cs_163bb9c6b14043fd3b40d34783681f41f380a650';
  var options = {
    'method': 'get',
    'headers': {
      'Authorization': 'Basic ' + Utilities.base64Encode(consumerKey + ':' + consumerSecret)
    }
  };
  
  // Add date filter and status filter to the API URL
  apiUrl += '?after=' + startDate.toISOString() + '&status=completed,credit';
  
  var page = 1;
  var orders = [];
  var fetchedOrders;
  
  do {
    var paginatedUrl = apiUrl + '&page=' + page + '&per_page=100'; // Fetch 100 orders per page
    var response = UrlFetchApp.fetch(paginatedUrl, options);
    fetchedOrders = JSON.parse(response.getContentText());
    orders = orders.concat(fetchedOrders);
    page++;
  } while (fetchedOrders.length > 0);
  
  // Get existing custom order numbers from the sheet
  var lastRow = sheet.getLastRow();
  var existingOrderNumbers = new Set();
  if (lastRow > 1) {
    existingOrderNumbers = new Set(sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat());
  }
  
  // Set headers if the sheet is empty
  if (lastRow === 0) {
    var headers = ['Invoice Date', 'Custom Order No', 'Customer Name', 'Product', 'Amount', 'Delivery Method', 'Discount', 'Shipping', 'Payment Method', 'Order Status', 'Notes', 'Metadata', 'Product SKU', 'Product Count', 'Refund Amount'];
    sheet.appendRow(headers);
  }
  
  // Sort orders by ascending order date
  orders.sort(function(a, b) {
    return new Date(a.date_created) - new Date(b.date_created);
  });
  
  // Remove existing orders with the same order number
  orders.forEach(function(order) {
    var customOrderNo = '';
    order.meta_data.forEach(function(meta) {
      if (meta.key === '_ywson_custom_number_order_complete') { // Adjust this key based on your plugin's metadata key
        customOrderNo = meta.value;
      }
    });
    
    // Remove existing order if it already exists
    if (existingOrderNumbers.has(customOrderNo)) {
      removeOrderByNumber(sheet, customOrderNo);
    }
  });
  
  // Append order data
  orders.forEach(function(order) {
    var customOrderNo = '';
    var refundAmount = 0;
    
    order.meta_data.forEach(function(meta) {
      if (meta.key === '_ywson_custom_number_order_complete') { // Adjust this key based on your plugin's metadata key
        customOrderNo = meta.value;
      }
    });
    
    // Check for refunds
    if (order.refunds && order.refunds.length > 0) {
      refundAmount = order.refunds.reduce((total, refund) => total + parseFloat(refund.total), 0);
    }
    
    Logger.log('Order ID: ' + order.id + ', Refund Amount: ' + refundAmount);
    
    var metadata = order.meta_data.map(meta => meta.key + ': ' + meta.value).join(', ');
    var productNames = [];
    var productSKUs = [];
    var productCounts = {};
    
    order.line_items.forEach(function(item) {
      if (!productSKUs.includes(item.sku)) {
        productSKUs.push(item.sku);
      }
      productNames.push(item.name);
      if (productCounts[item.name]) {
        productCounts[item.name]++;
      } else {
        productCounts[item.name] = 1;
      }
    });
    
    var row = [
      order.date_created,
      customOrderNo,
      order.billing.first_name + ' ' + order.billing.last_name,
      productNames.join(', '),
      order.total,
      order.shipping_lines.map(line => line.method_title).join(', '),
      order.discount_total,
      order.shipping_total,
      order.payment_method_title,
      order.status,
      order.customer_note,
      metadata,
      productSKUs.join(', '),
      Object.entries(productCounts).map(([name, count]) => `${name}: ${count}`).join(', '),
      refundAmount
    ];
    sheet.appendRow(row);
    
    // Add the custom order number to the set
    existingOrderNumbers.add(customOrderNo);
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