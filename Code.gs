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
