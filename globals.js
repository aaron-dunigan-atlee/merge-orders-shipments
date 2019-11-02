/**********************************
 * Global variables and constants *
 **********************************
 
 Variables we will use throughout the script. */

// Column: property maps
// How properties (headers) on the input sheets correspond to
// column numbers on the merged sheet.  Column numbers here are 0-indexed,
// i.e. 0=column A, 1=column B, etc.
var mergedSheetMainEntryOrderColumns = {
  0: 'orderKey',
  1: 'orderDate',
  2: 'orderNumber',
  3: 'shipTo_company',
  4: 'billTo_name',
  8: 'advancedOptions_storeId',
  17: 'paymentMethod',
  18: 'customerNotes',
  19: 'internalNotes',
  20: 'customerEmail',
  21: 'billTo_phone',
  22: 'billTo_street1',
  23: 'billTo_street2',
  24: 'billTo_city',
  25: 'billTo_state',
  26: 'billTo_country',
  27: 'billTo_postalCode',
  28: 'shipTo_street1',
  29: 'shipTo_street2',
  30: 'shipTo_city',
  31: 'shipTo_state',
  32: 'shipTo_country',
  33: 'shipTo_postalCode',
  34: 'requestedShippingService'
};

var mergedSheetItemOrderColumns = {
  0: 'orderKey',
  1: 'orderDate',
  2: 'orderNumber',
  10: 'items_name',
  11: 'items_sku',
  12: 'items_quantity',
  13: 'items_unitPrice',
  15: 'items_taxAmount',
  16: 'items_shippingAmount'
};

var mergedSheetItemShipmentColumns = {
  0: 'orderKey',
  2: 'orderNumber',
  35: 'orderId',
  38: 'shipDate',
  40: 'shipmentCost',
  41: 'insuranceCost',
  42: 'trackingNumber',
  47: 'insuranceOptions_insuredValue',
  48: 'shipmentItems_quantity',
  49: 'shipmentItems_name'
};

// Column letters 
// (for use in formula references).
var STORE_ID_COLUMN_LETTER = 'I';
var ITEM_PRICE_COLUMN_LETTER = 'N';
var ITEM_QTY_COLUMN_LETTER = 'M';
var ITEM_TOTAL_COLUMN_LETTER = 'O';
var SHIP_DATE_COLUMN_LETTER = 'AM';

// Column indices
// Indices for columns which will contain formulas or formatting.  
// As above, these are 0-based (Column A is 0, B is 1, etc.).
var STORE_NAME_COLUMN_INDEX = 7; // Column H
var DIMENSIONS_COLUMN_INDEX = 46; // Column AU
var WEIGHT_COLUMN_INDEX = 45; // Column AT
var CARRIER_COLUMN_INDEX = 36; // Column AK
var SERVICE_COLUMN_INDEX = 37; // Column AL
var CARRIER_USED_COLUMN_INDEX = 43; // Column AR
var SERVICE_USED_COLUMN_INDEX = 44; // Column AS
var ITEM_TOTAL_COLUMN_INDEX = 14; // Column O
var ITEM_SHIPPED_COLUMN_INDEX = 6; // Column G
var ITEM_NAME_COLUMN_INDEX = 10; // Column K
var ORDER_FULFILLED_COLUMN_INDEX = 5; // Column F
var ORDER_TOTAL_COLUMN_INDEX = 9; // Column J
var ORDER_DATE_COLUMN_INDEX = 1; // Column B
var QUARTER_COLUMN_INDEX = 39; // Column AN
var ORDER_KEY_COLUMN_INDEX = 0; // Column A

// This constant is needed because javascript arrays are zero-based (first item has index 0),
// whereas Google Sheets columns are indexed starting at 1.
var COLUMN_INDEX_OFFSET = 1;

// Array of merged data, before we write it to the sheet.
var MERGED_DATA = [];

// MAIN_ENTRY_ROWS will be filled with a list of A1-notation ranges based on whether 
// the corresponding row in MERGED_DATA is a main order entry or an individual item.
var MAIN_ENTRY_ROWS = [];

// Names of sheets from the spreadsheet file.
//var ORDERS_SHEET_NAME = "Orders";
//var SHIPMENTS_SHEET_NAME = "Shipments";
var MERGED_SHEET_NAME = "Orders and Shipments";

// ID's of external sheets.
var ORDERS_SHEET_ID = '14zfCISZvAfcYdLVYpOxphP3NpYSFqLt7ygPZnUzJYgQ';
var SHIPMENTS_SHEET_ID = '1b3MlpyA8D2xgdqzb3HLOOobLX6VG3eabE-pgR7kqMl0';

// Sheet objects
var MERGED_SHEET = SpreadsheetApp.getActive().getSheetByName(MERGED_SHEET_NAME);

// Old way of referencing orders and shipments.
//var ORDERS_SHEET = SpreadsheetApp.getActive().getSheetByName(ORDERS_SHEET_NAME);
//var SHIPMENTS_SHEET = SpreadsheetApp.getActive().getSheetByName(SHIPMENTS_SHEET_NAME);

// Width of merged sheet and a template of empty strings to start off each row
var MERGED_SHEET_WIDTH = MERGED_SHEET.getLastColumn();
var EMPTY_ROW = filledArray(MERGED_SHEET_WIDTH, "");

// Color to use when shading first row of each entry.
var SHADING_COLOR = 'Azure';

// How many header rows on the merged sheet. 
var MERGED_SHEET_HEADER_ROW_COUNT = 3;