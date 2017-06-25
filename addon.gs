/**
 * Holds unique vendor information for Indira Active's suppliers.
 * @param {String} name References unofficial internal name
 * @param {String} spreadsheetID Google Spreadsheet id from vendor product mapping & inventory
 * @param {String} sheetName Google Sheet name with current inventory 
 * @param {String} inventoryEmail Sparkshipping import address for vendor.
 */
function vendor(name, spreadsheetID, sheetName, inventoryEmail) {
    this.name = name;
    this.spreadsheetID = spreadsheetID;
    this.sheetName = sheetName;
    this.inventoryEmail = inventoryEmail;
}


/**
 * Initalizes config and vendors.
 */
function init() {
    config = {
      url_ext: 'exportFormat=csv&format=csv' // export as pdf / csv / xls / xlsx
      +
      '&size=letter' // paper size legal / letter / A4
      +
      '&portrait=false' // orientation, false for landscape
      +
      '&fitw=true&source=labnol' // fit to page width, false for actual size
      +
      '&sheetnames=false&printtitle=false' // hide optional headers and footers
      +
      '&pagenumbers=false&gridlines=false' // hide page numbers and gridlines
      +
      '&fzr=false' // do not repeat row headers (frozen rows) on each page
      +
      '&gid=', // the sheet's Id
      currentTime: new Date().toLocaleString(),
      folder: "",
      currentSpreadsheet: ""
    };  
  
    vendors = [];
    // Initialize Suppliers
    var ardenFulfillment = new vendor('Arden Fulfillment', 'UNIQUE_SPREADSHEET_ID', 'Internal Inventory', 'INVENTORY_SYSTEM_EMAIL_ADDRESS');
    vendors.push(ardenFulfillment);
}


/**
 * Retrieves and formats inventory from Google Sheet.
 * @param {Object} vendor Holds unique vendor information
 * @returns {Object} blob Binary CSV export of spreadsheet
 */
function prepData(vendor) {
    var ss = SpreadsheetApp.openById(vendor.spreadsheetID);
    Logger.log(ss.getName());
    var sheet = ss.getSheetByName(vendor.sheetName);
    var url = "https://docs.google.com/spreadsheets/d/SS_ID/export?".replace("SS_ID", ss.getId());
    var token = ScriptApp.getOAuthToken();

    // Fetch Spreadsheet
    var response = UrlFetchApp.fetch(url + config.url_ext + sheet.getSheetId(), {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    // Format blob data to CSV and save in Drive
    var blob = response.getBlob().setName(vendor.name + ' inventory as of ' + config.currentTime + '.csv');
    var file = config.folder.createFile(blob);

    return blob;
}


/**
 * Sends inventory levels to inventory system.
 * @param {Object} vendor Holds unique vendor information
 * @param {Object} blob Binary CSV export of spreadsheet
 */
function syncInventory(vendor, blob) {
    var subject = vendor.name + " inventory as of " + config.currentTime;
    var body = 'Automatic inventory import to Sparkshipping.';
    if (MailApp.getRemainingDailyQuota() > 0)
        GmailApp.sendEmail(vendor.inventoryEmail, subject, body, {
            attachments: [blob],
            noReply: true,
            name: "Indira Active - Inventory",
            bcc: "archive@example.com"
        });
    Logger.log("Sent email to " + vendor.inventoryEmail + " for " + vendor.name + ".");
}


/**
 * Prevents user from running syncAllVendors before previous finishes.
 */
function waitForScript() {
  Utilities.sleep(60000);
}


/**
 * Runs when the addon is installed.
 * @param {Object} e Contextual trigger event information
 */
function onInstall(e) {
  Logger.log("onInstall()");
  init();
  onOpen(e);
}


/**
 * Runs when the document is loaded (and thus addon initalized).
 * @param {Object} e Contextual trigger event information
 */
function onOpen(e) {
  init();
  Logger.log("onOpen()");
  var menu = SpreadsheetApp.getUi().createAddonMenu(); 
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    // Add a normal menu item (works in all authorization modes).
    menu.addItem('Sync current supplier (recommended)', 'syncCurrentVendor');
    menu.addItem('Sync all suppliers', 'syncAllVendors');
  } else {
    // Add a menu item based on properties (doesn't work in AuthMode.NONE).
    var properties = PropertiesService.getDocumentProperties();
    var workflowStarted = properties.getProperty('workflowStarted');
    if (workflowStarted) {
      menu.addItem('Sync in progress... ', 'waitForScript');
    } else {
      menu.addItem('Sync current supplier (recommended)', 'syncCurrentVendor');
      menu.addItem('Sync all suppliers', 'syncAllVendors');
    }
  }
  menu.addToUi();
}


/**
 * Syncs inventory for all suppliers.
 */
function syncAllVendors() {
  init();
  var dir = DriveApp.getFolderById("0Bwa8AXm6TbI4WDU0MHNCRW9TWms"); 
  config.folder = dir.createFolder('Inventory as of ' + config.currentTime);
  
  for (var i = 0; i < vendors.length; i++) {
    inventory = prepData(vendors[i]);
    syncInventory(vendors[i], inventory);
  }
}


/**
 * Syncs inventory for current spreadsheet.
 */
function syncCurrentVendor() {
  init();
  config.currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet().getId();
  var dir = DriveApp.getFolderById("UNIQUE_DRIVE_FOLDER_ID"); 
  config.folder = dir.createFolder('Inventory as of ' + config.currentTime);
  
  for (var i = 0; i < vendors.length; i++) {
    Logger.log(config.currentSpreadsheet)
    Logger.log(vendors[i].spreadsheetID)
    if (config.currentSpreadsheet == vendors[i].spreadsheetID) {
      inventory = prepData(vendors[i]);
      syncInventory(vendors[i], inventory);
      break;
    }
  }
}