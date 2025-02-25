function fetchWooCommerceData() {
  var sheetName = 'Woocommerce Daily Data'; // Update this to match your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }
  
  // Calculate the start and end of the current day
  var now = new Date();
  var startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  var endDate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  
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
  apiUrl += '?after=' + startDate.toISOString() + '&before=' + endDate.toISOString() + '&status=completed,credit';
  
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
  
  // Clear existing data
  sheet.clearContents();
  
  // Set headers
  var headers = ['Invoice Date', 'Custom Order No', 'Customer Name', 'Product', 'Amount', 'Delivery Method', 'Discount', 'Shipping', 'Payment Method', 'Order Status', 'Notes', 'Metadata', 'Product SKU', 'Product Count'];
  sheet.appendRow(headers);
  
  // Sort orders by descending order date
  orders.sort(function(a, b) {
    return new Date(b.date_created) - new Date(a.date_created);
  });
  
  // Append order data
  orders.forEach(function(order) {
    var customOrderNo = '';
    order.meta_data.forEach(function(meta) {
      if (meta.key === '_ywson_custom_number_order_complete') { // Adjust this key based on your plugin's metadata key
        customOrderNo = meta.value;
      }
    });
    
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
    
    var customerName = order.billing.first_name + ' ' + order.billing.last_name;
    if (customerName.trim() === '') {
      customerName = 'Guest';
    }
    
    var row = [
      order.date_created,
      customOrderNo,
      customerName,
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
      Object.entries(productCounts).map(([name, count]) => `${name}: ${count}`).join(', ')
    ];
    sheet.appendRow(row);
  });
}