function fetchWooCommerceData() {
  var sheetName = 'Restock and Cost Data'; // Update this to match your sheet name
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet not found: ' + sheetName);
    return;
  }
  
  // Calculate the start and end of the last three days in UTC
  var now = new Date();
  var startDate = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate() - 2));
  var endDate = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate() + 1));
  
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
  
  // Sort orders by date in descending order
  orders.sort(function(a, b) {
    return new Date(b.date_created) - new Date(a.date_created);
  });
  
  // Clear existing data
  sheet.clearContents();
  
  // Set headers
  var headers = ['Order No', 'Product Name', 'Product Count', 'Product Price', 'Product Cost', 'Product Metadata', 'Order Discount', 'Order Date'];
  sheet.appendRow(headers);
  
  // Append order data
  orders.forEach(function(order) {
    var orderDiscount = order.discount_total ? parseFloat(order.discount_total) : 0;
    var orderDate = new Date(order.date_created).toLocaleDateString(); // Get the order date
    order.line_items.forEach(function(item) {
      try {
        // Fetch the product details using the product ID
        var productUrl = 'https://nirasu.com/wp-json/wc/v3/products/' + item.product_id;
        var productResponse = UrlFetchApp.fetch(productUrl, options);
        var product = JSON.parse(productResponse.getContentText());
        
        // Get the product cost
        var productCost = 'N/A';
        var productMetadata = '';
        var productPrice = product.price; // Default to product price
        
        if (item.variation_id) {
          // Fetch the variant details
          var variantUrl = 'https://nirasu.com/wp-json/wc/v3/products/' + item.product_id + '/variations/' + item.variation_id;
          var variantResponse = UrlFetchApp.fetch(variantUrl, options);
          var variant = JSON.parse(variantResponse.getContentText());
          var variantCostMeta = variant.meta_data.find(meta => meta.key === '_alg_wc_cog_cost');
          productCost = variantCostMeta ? variantCostMeta.value : 'N/A';
          productMetadata = variant.meta_data.map(meta => meta.key + ': ' + meta.value).join(', ');
          productPrice = variant.price; // Use variant price if available
        } else {
          var productCostMeta = product.meta_data.find(meta => meta.key === '_alg_wc_cog_cost');
          productCost = productCostMeta ? productCostMeta.value : 'N/A';
          productMetadata = product.meta_data.map(meta => meta.key + ': ' + meta.value).join(', ');
        }
        
        var row = [
          order.id,
          item.name,
          item.quantity, // Get the product count
          productPrice, // Get the product price, considering variants
          productCost, // Get the product cost from metadata
          productMetadata, // Get the product metadata
          orderDiscount, // Get the order discount
          orderDate // Add the order date as the last column
        ];
        sheet.appendRow(row);
      } catch (e) {
        Logger.log('Error fetching product details for product ID: ' + item.product_id + ' - ' + e.message);
      }
    });
  });
}