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
  var nameCol = rangeValues[0].indexOf("Name");
  var numberCol = rangeValues[0].indexOf("Number");  
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
  
  //  Add Contact Group header with hyperlink
  urlArray.push(['=HYPERLINK("https://contacts.google.com/label/' + groupID + '", "Contact Group")']);
  
  //  Get contacts from sheet and save to Google account as Contacts  
  for (var x = headerRow; x <= rangeValues.length - 1; x++){  
    
    //    Set first and last names and phone number (with dashes)
    firstName = rangeValues[x][nameCol].toString().split(" ")[0];
    lastName = rangeValues[x][nameCol].toString().split(" ")[rangeValues[x][nameCol].toString().split(" ").length - 1];
    phoneNumber = rangeValues[x][numberCol].replace(/ /gi, "-");
    
    //    Create contact with names and number
    console.log(firstName + " " + lastName);
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
  
  //  Print saved links to Google Sheet
  sheet.getRange(1, sheetRange.getLastColumn() + 1, urlArray.length, 1).setValues(urlArray);
  
}
