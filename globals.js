/**********************************
 * Global variables and constants *
 **********************************
 
 Variables we will use throughout the script. */

// Column letters 
// (for use in formula references).
var ITEM_PRICE_COLUMN_LETTER = 'N';
var ITEM_QTY_COLUMN_LETTER = 'M';
var ITEM_TOTAL_COLUMN_LETTER = 'O';
var SHIP_DATE_COLUMN_LETTER = 'AM';

// Column indices
// Indices for columns which will contain formulas or formatting.  
// As above, these are 0-based (Column A is 0, B is 1, etc.).
var STORE_NAME_COLUMN_INDEX = getColumnIndex('merged_storeName');
var STORE_ID_COLUMN_INDEX = getColumnIndex('orders_advancedOptions_storeId');
var DIMENSIONS_COLUMN_INDEX = getColumnIndex('merged_dimensions');
var WEIGHT_COLUMN_INDEX = getColumnIndex('merged_weight');
var CARRIER_COLUMN_INDEX = getColumnIndex('merged_carrierCode');
var SERVICE_COLUMN_INDEX = getColumnIndex('merged_serviceCode');
var CARRIER_USED_COLUMN_INDEX = getColumnIndex('merged_carrierUsed');
var SERVICE_USED_COLUMN_INDEX = getColumnIndex('merged_serviceUsed');
var ITEM_TOTAL_COLUMN_INDEX = getColumnIndex('merged_itemTotal');
var ITEM_PRICE_COLUMN_INDEX = getColumnIndex('orders_items_unitPrice');
var ITEM_QTY_COLUMN_INDEX = getColumnIndex('orders_items_quantity');
var SHIP_DATE_COLUMN_INDEX = getColumnIndex('shipments_shipDate');
var ITEM_SHIPPED_COLUMN_INDEX = getColumnIndex('merged_shipped');
var ITEM_NAME_COLUMN_INDEX = getColumnIndex('orders_items_name');
var ORDER_FULFILLED_COLUMN_INDEX = getColumnIndex('merged_fulfilled');
var ORDER_TOTAL_COLUMN_INDEX = getColumnIndex('merged_orderTotal');
var ORDER_DATE_COLUMN_INDEX = getColumnIndex('orders_orderDate');
var QUARTER_COLUMN_INDEX = getColumnIndex('merged_quarter');
var ORDER_KEY_COLUMN_INDEX = getColumnIndex('orders_orderKey');
var IS_HEADER_COLUMN_INDEX = getColumnIndex('merged_orderHeader');

// This constant is needed because javascript arrays are zero-based (first item has index 0),
// whereas Google Sheets columns are indexed starting at 1.
var COLUMN_INDEX_OFFSET = 1;
var ROW_INDEX_OFFSET = 1;

// Global variables for holding object data, before/after we write it to the sheet.
var mergedDataToUpdate = {};
var mergedDataToAdd = {};
var ordersData = {};
var shipmentsData = {};
var existingMergedData = {};

// MAIN_ENTRY_ROWS will be filled with a list of A1-notation ranges based on whether 
// the corresponding row in mergedDataToUpdate is a main order entry or an individual item.
var MAIN_ENTRY_ROWS = [];

// Names of sheets from the spreadsheet file.
var MERGED_SHEET_NAME = "Orders and Shipments";

// Sheet objects
var MERGED_SHEET = SpreadsheetApp.getActive().getSheetByName(MERGED_SHEET_NAME);

// Header names in merged sheet.
var MERGED_SHEET_HEADERS = MergeDb.getHeaders(MERGED_SHEET_NAME);

// Width of merged sheet and a template of empty strings to start off each row
var MERGED_SHEET_WIDTH = MERGED_SHEET.getLastColumn();
var EMPTY_ROW = filledArray(MERGED_SHEET_WIDTH, "");

// Color to use when shading first row of each entry.
var SHADING_COLOR = 'Azure';

// How many header rows on the merged sheet. 
var MERGED_SHEET_HEADER_ROW_COUNT = 4;

// Property names that go into header rows for each order.
var ORDER_HEADER_FIELDS = getOrderHeaderFields();

