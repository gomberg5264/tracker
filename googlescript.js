// This is Google Apps Script - meant only to show code

function onFormSubmit(e) { // triggered function on submit button 
    var uniqueID = getUniqueID(e.values); // function to create a uniqueID
    recordResponseID(e.range, uniqueID);
    sendAutomatedEmail(); // function to send email on form submit
  }
  
  // records uniqueID to correct cell
  function recordResponseID(eventRange, uniqueID) { 
    var row = eventRange.getLastRow(); // param {Object} eventRange range where response is recorded
    var sheet = SpreadsheetApp.getActiveSheet(); // param {Integer} uniqueID for range
    sheet.getRange(row, 1).setValue(uniqueID);
  
  }
  
  // function to get form connected to spreadsheet
  function getConnectedForm() {
    var formUrl = SpreadsheetApp.getActiveSpreadsheet().getFormUrl();
    var form =  FormApp.openByUrl(formUrl);
    return form;
  }
  
  // returns uniqueID for response
  function getUniqueID(eventValues) {
    var isMatch = false;
    var eventItems = eventValues.slice(1);
  
    var responses = getConnectedForm().getResponses();
    //loop backwards through responses (latest is most likely)
    for (var i = responses.length - 1; i > -1; i--) {
      var responseItems = responses[i].getItemResponses();
      //check each value matches
  
      for (var j = 0; j < responseItems.length; j++) {
        if (responseItems[j].getResponse() !== eventItems[j]) {
          break;
        }
        isMatch = true;
      }
      if (isMatch) {
        return i + 1;
      }
    }
  }
  
  // automated email when a customer submits an order
  function sendAutomatedEmail() {
    
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1").activate(); 
    var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();  // Get main Database Sheet
    var emailsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Email Template");  // Get Sheet with the Email Content
    var lrow = sh.getLastRow();   // User which last filled the form
    
    var id = sh.getRange(lrow, 1).getValue();     // Get id value. Row - Last Row, Column - 1
    var name = sh.getRange(lrow, 3).getValue();     // Get name value. Row - Last Row, Column - 3
    var email = sh.getRange(lrow, 4).getValue();    // Get email value. Row - Last Row, Column - 4
    var status = sh.getRange(lrow, 5).setValue("New");    // Set Status value to New. Row - Last Row, Column - 5
    var subject = "Order Recieved";
    
    var header = emailsheet.getRange(1 , 1).getValue();   // Get Header of email
    header = header.replace("{name}", name);   // Replace variable 'name' with user's name
    var message = emailsheet.getRange(2 , 1).getValue();   // Get Email Message
    message = message.replace("{id}", id);  // Replace variable 'message' with email message
    
    var body = header + message; // body of message to be with header and text
    Logger.log(body);
    
    MailApp.sendEmail(email, subject, body);    // Send email to email entered by the user
    Logger.log(body);
  }
  
  /// Status change functions
  var admin_email='nohxx044@umn.edu'; //<- static email address to be shipped
  
  function triggerOnEdit(e) // function that gets trigger when status is changed to Ready to ship
  {
    sendEmailOnApproval(e);
  }
  
  // function for when status changes to "shipped"
  function anotherOnEdit(e)
  {
    sendEmailOnShipped(e); // sends out shipped email
    shippedOrders(); // moves orders with a status of shipped to a new tab
  }
  
  function checkStatusIsApproved(e) // function to check the status
  {
    var range = e.range;
    
    if(range.getColumn() <= 5 && // row 5 is the status
       range.getLastColumn() >=5 )
    {
      var edited_row = range.getRow();
      
      var status = SpreadsheetApp.getActiveSheet().getRange(edited_row,5).getValue();
      if(status == 'Ready to be Shipped') // when status is "Ready to be Shipped" triggers sendEmailOnApproval
      {
        return edited_row;
      }
    return 0;
  }
  }
  
  function checkStatusIsShipped(e) // function to check for shipped
  {
    var range = e. range;
    
    if(range.getColumn() <= 5 &&
       range.getLastColumn() >=5)
    {
      var target_row = range.getRow();
      
      var status = SpreadsheetApp.getActiveSheet().getRange(target_row,5).getValue();
      if(status == 'Shipped')
      {
        return target_row;
      }
      return 0;
    }
  }
  
  function sendEmailOnShipped(e) // email on shipped
  {
    var shipped_row = checkStatusIsShipped(e);
    
    if(shipped_row <= 0)
    {
      return;
    }
    sendEmailShipped(shipped_row);
  }
  
  function sendEmailShipped(row)
  {
    var values = SpreadsheetApp.getActiveSheet().getRange(row,1,row,5).getValues();
    var row_values = values[0];
    
    var customer_email = composeCustomerEmail(row_values);
    
    SpreadsheetApp.getUi().alert(" CUSTOMER TEST subject is "+customer_email.subject+"\n message "+customer_email.message); // to test message of email
  
    MailApp.sendEmail(customer_email.email,customer_email.subject,customer_email.message);
  }
    
    
   function sendEmailOnApproval(e) // when status is 'ready to be shipped', sends an email to the admin
  {
    var approved_row = checkStatusIsApproved(e);
    
    if(approved_row <= 0)
    {
      return;
    }
    sendEmailByRow(approved_row);
  }
  
  function sendEmailByRow(row)
  {
    var values = SpreadsheetApp.getActiveSheet().getRange(row,1,row,5).getValues();
    var row_values = values[0];
    
    var mail = composeApprovedEmail(row_values);
    
    SpreadsheetApp.getUi().alert("To Admin: subject is "+mail.subject+"\n message "+mail.message); // to test message of email
    
    MailApp.sendEmail(admin_email,mail.subject,mail.message);
    
  }
  
  // composes email to admin
  function composeApprovedEmail(row_values) 
  {
    var orderNumber= row_values[0];
    var name = row_values[2];
    var email = row_values[3];
  
    var message = "The following applicant's order is ready to be shipped: "+name+" email "+ email +" order number " + orderNumber;
    var subject = "Order is Ready to be shipped for order " + orderNumber;
    
    return({message:message,subject:subject});
  }
  
  function composeCustomerEmail(row_values) // email for customer when order status has changed to "Shipped"
  {
    var orderNumber = row_values[0];
    
    var name = row_values[2];
    
    var email = row_values[3];
  
    var message = "Hello "+name+" your order # "+orderNumber+ " has been shipped out! You can except your order in 3-5 days.";
    var subject = name+ ", Your Order " +orderNumber+ "has been shipped!";
    
    return({message:message,subject:subject, email:email });
  }
  
  // function to move shipped orders to a new tab
  
  function shippedOrders() {
    // moves a row from a sheet to another when status is shipped
    var sheetNameToWatch = "Form Responses 1";
    var columnNumberToWatch = 5; // Status Column
    var valueToWatch = "Shipped";
    var sheetNameToMoveTheRowTo = "Shipped Orders";
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getActiveCell();
  
    if (sheet.getName() == sheetNameToWatch && range.getColumn() == columnNumberToWatch && range.getValue() == valueToWatch) {
  
      var targetSheet = ss.getSheetByName(sheetNameToMoveTheRowTo);
      var targetRange = targetSheet.getRange(targetSheet.getLastRow() + 1, 1);
      sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn()).copyTo(targetRange);
      sheet.deleteRow(range.getRow());
    }
  }
  
  