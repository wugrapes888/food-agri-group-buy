// ============================================================
// 團購發貨系統 - 主入口
// 部署前請在 Script Properties 設定：GEMINI_API_KEY
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('團購發貨系統');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- 前端呼叫的 API ---

function getProducts() {
  return SheetDB.getProducts();
}

function getAllCustomers() {
  return SheetDB.getAllCustomers();
}

function getOrderSummary() {
  return SheetDB.getOrderSummary();
}

function getOrdersForExport(productNameOrGroup) {
  return SheetDB.getOrdersForExport(productNameOrGroup);
}

function createExportFile(rows, title) {
  const ss = SpreadsheetApp.create(title || '訂單匯出');
  const sheet = ss.getActiveSheet();
  sheet.setName('訂單');
  sheet.getRange(1, 1, 1, 6).setValues([['客人姓名','商品','數量','單價','小計','取貨狀態']]);
  sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  if (rows.length > 0) {
    const data = rows.map(r => [r.customer, r.product, r.qty, r.price||0, r.subtotal||0, r.status]);
    sheet.getRange(2, 1, data.length, 6).setValues(data);

    // ── 品項彙整區 ──
    const summaryStartRow = data.length + 3;

    // 標題
    sheet.getRange(summaryStartRow, 1, 1, 4).setValues([['品項彙整', '', '', '']]);
    sheet.getRange(summaryStartRow, 1).setFontWeight('bold').setFontSize(12);

    // 小標題
    sheet.getRange(summaryStartRow + 1, 1, 1, 4).setValues([['商品', '訂購總數', '單價', '小計']]);
    sheet.getRange(summaryStartRow + 1, 1, 1, 4).setFontWeight('bold').setBackground('#e8eaf6');

    // 統計每個品項
    const productMap = {};
    const productOrder = [];
    rows.forEach(r => {
      if (!productMap[r.product]) {
        productMap[r.product] = { qty: 0, price: r.price || 0 };
        productOrder.push(r.product);
      }
      productMap[r.product].qty += r.qty;
    });

    const summaryData = productOrder.map(name => {
      const p = productMap[name];
      return [name, p.qty, p.price, p.qty * p.price];
    });
    sheet.getRange(summaryStartRow + 2, 1, summaryData.length, 4).setValues(summaryData);

    // 合計列
    const grandTotal = summaryData.reduce((s, r) => s + r[3], 0);
    const totalQty = summaryData.reduce((s, r) => s + r[1], 0);
    const totalRow = summaryStartRow + 2 + summaryData.length;
    sheet.getRange(totalRow, 1, 1, 4).setValues([['合計', totalQty, '', grandTotal]]);
    sheet.getRange(totalRow, 1, 1, 4).setFontWeight('bold').setBackground('#c5cae9');
  }
  sheet.autoResizeColumns(1, 6);
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
}

function getBuyersByProduct(productName) {
  return SheetDB.getBuyersByProduct(productName);
}

function getCustomerDetail(customerName) {
  return SheetDB.getCustomerDetail(customerName);
}

function completePickup(customerName) {
  return SheetDB.completePickup(customerName);
}

function undoPickup(customerName) {
  return SheetDB.undoPickup(customerName);
}

function parseLineScreenshot(base64Image) {
  return OCR.parseImage(base64Image);
}

function importOrders(orders) {
  return SheetDB.importOrders(orders);
}

function clearAllData() {
  return SheetDB.clearAllData();
}

function getStats() {
  return SheetDB.getStats();
}

function getProductsForManagement() {
  return SheetDB.getProductsForManagement();
}

function saveProduct(product) {
  return SheetDB.saveProduct(product);
}

function deleteProduct(name) {
  return SheetDB.deleteProduct(name);
}

function batchSaveProducts(products) {
  return SheetDB.batchSaveProducts(products);
}

function parseTextOrders(text) {
  return OCR.parseText(text);
}

function deleteOrderRow(customerName, productName) {
  return SheetDB.deleteOrderRow(customerName, productName);
}

function deleteCustomerAllOrders(customerName) {
  return SheetDB.deleteCustomerAllOrders(customerName);
}

function setupApiKey() {
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', 'AIzaSyChXpE3FYQdMgwy4x3eabMla1sNaGF_Si0');
  return 'API Key 設定完成';
}
