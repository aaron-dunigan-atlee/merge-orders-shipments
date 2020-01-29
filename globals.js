/**********************************
 * Global variables and constants *
 **********************************
 
 Variables we will use throughout the script. */

// This constant is needed because javascript arrays are zero-based (first item has index 0),
// whereas Google Sheets columns are indexed starting at 1.
var COLUMN_INDEX_OFFSET = 1;
var ROW_INDEX_OFFSET = 1;

// How many header rows on the merged sheet. 
var MERGED_SHEET_HEADER_ROW_COUNT = 5;

// Global variables for holding object data, before/after we write it to the sheet.
var mergedDataToUpdate = {};
var mergedDataToAdd = {};
var ordersData = {};
var shipmentsData = {};
var existingMergedData = {};

// Names of sheets from the spreadsheet file.
var MERGED_SHEET_NAME = "Orders and Shipments";

// Sheet objects
var MERGED_SHEET = SpreadsheetApp.getActive().getSheetByName(MERGED_SHEET_NAME);

// Header names in merged sheet.
var MERGED_SHEET_HEADERS = [];

// Width of merged sheet and a template of empty strings to start off each row
var MERGED_SHEET_WIDTH = MERGED_SHEET.getLastColumn();
var EMPTY_ROW = filledArray(MERGED_SHEET_WIDTH, "");


// Property names that go into header rows for each order.
var ORDER_HEADER_FIELDS = getOrderHeaderFields();

// Indices for columns which will contain formulas or formatting.  
// As above, these are 0-based (Column A is 0, B is 1, etc.).
var STORE_NAME_COLUMN_INDEX;
var STORE_ID_COLUMN_INDEX;
var DIMENSIONS_COLUMN_INDEX;
var WEIGHT_COLUMN_INDEX;
var CARRIER_COLUMN_INDEX;
var SERVICE_COLUMN_INDEX;
var CARRIER_USED_COLUMN_INDEX;
var SERVICE_USED_COLUMN_INDEX;
var ITEM_TOTAL_COLUMN_INDEX;
var ITEM_PRICE_COLUMN_INDEX;
var ITEM_QTY_COLUMN_INDEX;
var SHIP_DATE_COLUMN_INDEX;
var ITEM_SHIPPED_COLUMN_INDEX;
var ITEM_NAME_COLUMN_INDEX;
var ORDER_FULFILLED_COLUMN_INDEX;
var ORDER_TOTAL_COLUMN_INDEX;
var ORDER_DATE_COLUMN_INDEX;
var QUARTER_COLUMN_INDEX;
var ORDER_KEY_COLUMN_INDEX;
var IS_HEADER_COLUMN_INDEX;
var MAIN_ENTRY_PROPERTIES;
var MERGED_SHEET_TITLES;

// Set values for some of these global variables.
// For technical reasons, tehse must be initialized inside a function.
function initializeGlobals() {
  // Initialize global variables.
  mergedDataToUpdate = {};
  mergedDataToAdd = {};
  ordersData = getOrdersData();
  //Logger.log(ordersData);
  shipmentsData = getShipmentsData();
  //Logger.log(shipmentsData);
  console.time('get existingMergedData')
  existingMergedData = MergeDb.getJson(MERGED_SHEET);
  console.timeEnd('get existingMergedData')
  //Logger.log(existingMergedData);
  MERGED_SHEET_HEADERS = MergeDb.getHeaders(MERGED_SHEET_NAME);
  MAIN_ENTRY_PROPERTIES = MergeDb.getMainEntryProperties(MERGED_SHEET);

  STORE_NAME_COLUMN_INDEX = getColumnIndex('merged_storeName');
  STORE_ID_COLUMN_INDEX = getColumnIndex('orders_advancedOptions_storeId');
  DIMENSIONS_COLUMN_INDEX = getColumnIndex('merged_dimensions');
  WEIGHT_COLUMN_INDEX = getColumnIndex('merged_weight');
  CARRIER_COLUMN_INDEX = getColumnIndex('merged_carrierCode');
  SERVICE_COLUMN_INDEX = getColumnIndex('merged_serviceCode');
  CARRIER_USED_COLUMN_INDEX = getColumnIndex('merged_carrierUsed');
  SERVICE_USED_COLUMN_INDEX = getColumnIndex('merged_serviceUsed');
  ITEM_TOTAL_COLUMN_INDEX = getColumnIndex('merged_itemTotal');
  ITEM_PRICE_COLUMN_INDEX = getColumnIndex('orders_items_1_unitPrice');
  ITEM_QTY_COLUMN_INDEX = getColumnIndex('orders_items_1_quantity');
  SHIP_DATE_COLUMN_INDEX = getColumnIndex('shipments_shipDate');
  ITEM_SHIPPED_COLUMN_INDEX = getColumnIndex('merged_shipped');
  ITEM_NAME_COLUMN_INDEX = getColumnIndex('orders_items_1_name');
  ORDER_FULFILLED_COLUMN_INDEX = getColumnIndex('merged_fulfilled');
  ORDER_TOTAL_COLUMN_INDEX = getColumnIndex('merged_orderTotal');
  ORDER_DATE_COLUMN_INDEX = getColumnIndex('orders_orderDate');
  QUARTER_COLUMN_INDEX = getColumnIndex('merged_quarter');
  ORDER_KEY_COLUMN_INDEX = getColumnIndex('orders_orderKey');
  IS_HEADER_COLUMN_INDEX = getColumnIndex('merged_orderHeader');

}