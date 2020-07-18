/**
*
* Create contacts from Google Sheets (Col1 full of First and Last names, Col2 with numbers). Adds links to ColC.
*
* References
* https://developers.google.com/apps-script/reference/contacts
*
*/

function createContacts() {
  
  // Declare/initialize variables
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var sheetRange = sheet.getDataRange();
  var rangeValues = sheetRange.getDisplayValues();  
  var rangeFormulas = sheetRange.getFormulas();  
  var nameCol = rangeValues[0].indexOf("Name");
  var numberCol = rangeValues[0].indexOf("Number");  
  var linkCol = rangeValues[0].indexOf("Contact Group");
  var contact = "";
  var firstName = "";
  var lastName = "";
  var phoneNumber = "";
  var headerRow = 1;
  var urlArray = [];
  var splitContact = "";
  var contactID = "";
  var contactIDNumber = "";
  var contactbyID = "";  
  
  //  Create new Contact Group or get Contact Group by the name, add link
  var groupName = "Volunteer";
  var group = ContactsApp.getContactGroup(groupName) || ContactsApp.createContactGroup(groupName);
  var getGroupURL = group.getId();  
  var splitGroup = getGroupURL.split("/");
  var groupID = splitGroup[splitGroup.length - 1];
  
  //  Add Contact Group header with hyperlink if not already added
  urlArray.push(['=HYPERLINK("https://contacts.google.com/label/' + groupID + '", "Contact Group")']);
  
  //  Get contacts from sheet and save to Google account as Contacts  
  for (var x = headerRow; x <= rangeValues.length - 1; x++){  
    
    //    If there is a link to this contact already in ColC, skip the process by adding URL to array at end
    if (rangeValues[x][linkCol] != ""){
      urlArray.push([rangeFormulas[x][linkCol]]);
    } else {
      //    Set first and last names and phone number (with dashes)
      firstName = rangeValues[x][nameCol].toString().split(" ")[0];
      lastName = rangeValues[x][nameCol].toString().split(" ")[rangeValues[x][nameCol].toString().split(" ").length - 1];
      phoneNumber = rangeValues[x][numberCol].replace(/ /gi, "-");
      
      //    Create contact with names and number. Sleep script at several times to avoid error:
      //    "TypeError: Cannot read property 'addPhone' of null (line 53, file "Code"). The resource you requested could not be located."
      contact = ContactsApp.createContact(firstName, lastName, "");
      Utilities.sleep(5000);
      contactID = contact.getId();
      Utilities.sleep(5000);
      contactbyID = ContactsApp.getContactById(contactID);
      Utilities.sleep(5000);    
      contactbyID.addPhone(ContactsApp.Field.MAIN_PHONE, phoneNumber);
      Utilities.sleep(5000);
      
      //    Add to group
      group.addContact(contactbyID);
      
      //    Save Contact info
      splitContact = contactID.split("/");
      contactIDNumber = splitContact[splitContact.length - 1];
      urlArray.push(['=HYPERLINK("https://contacts.google.com/contact/' + contactIDNumber + '", "Contact Link for ' + rangeValues[x][nameCol] + '")']);
    }
  }
  
  //  Print saved links to Google Sheet
  sheet.getRange(1, sheetRange.getLastColumn(), urlArray.length, 1).setValues(urlArray);  
}

// ***********************************************************************************************************************

/**
*
* Create a menu option for script functions.
*
*/

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
  .addItem('Create Contacts from Sheet for Volunteer Group', 'createContacts')
  .addToUi();
}