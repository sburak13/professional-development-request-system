/**
 Google Forms Response to Google Docs Conversion
 
 Version 2.0
 Dillon Shu, Samantha Burak
 8/1/2019
 
 Trigger set to go off on form submit
 
 Adds a comments box to the end of the doc with the name of each person that needs to sign off on the professional development experience
 Shares the doc with those people (without the Google-generated message) and sends an email to them with the link to the document and instructions
 
 Test History:
 Add a new question to the template anywhere and to the form anywhere
 Add a new question to the template inside a second table inside the first one
 
 --------------------------------------------------------------------------------------------------------------------------------------------------------
 Version 1.0
 Dillon Shu, Samantha Burak, Aditya Gollapudi
 7/1/2019
 
 Title Format - last name , first name - name of thing they are going to, start date - end date
 
 Make copy of template document and put it in the Template Responses folder
 Find name and put it in {{name}}
 Search template for {{name}} and replace with form response
 Hyperlink for attached files
 
 Test History: 
 Weird Unicode 
 Emojis
 Other Languages
 Weird Files
 Multiple Files
 Invalid Dates
**/

// Global vars
var approvalRequestTemplateDocID = '1iqXCDi6IBix6PmvCO-tzR1eYuXP3lmC2xFcxb5w8X5E'
var pendingReponsesFolderID = '1FQajZxxrfNM6Znt_X49da4qlgVeEHNiY'
var facultyTableSheetURL = 'https://docs.google.com/spreadsheets/d/127XQ0elfDQiyH1wjgoltNilRROAO3U_c2bSqVoJS7m4/edit#gid=0'
var commentsBoxTemplateDocID = '1S6Ra71aTZqaqLs2W7WY5H3kKoZ9SRX24_uSaPyRIfg4'

var keywordSearchTerm = 'Person'
var attachFilesSeachTerm = 'Attach Files'

// Main function of this program
function grabResponse(e) {
  var docID = ""
  try 
  {
    // Gets the trigger object (in this case will trigger with each submission)
    var formResponse = e.response
    
    // Gets an array of the responses for each question from the submission
    var itemResponses = formResponse.getItemResponses()
    
    // Date Validation - sends email to respondent if end date is before start date and returns to end the program (careful of this when testing)
    // Uses helper function to return response to a specific question
    var startDate = getItemResponse("Start Date", itemResponses) 
    var endDate = getItemResponse("End Date", itemResponses)
    if (startDate > endDate) {
      sendInvalidDatesErrorEmail(formResponse.getRespondentEmail())
      return
    }
    
    // Makes a copy of the template document using its ID and copies it to the target folder
    var template = DriveApp.getFileById(approvalRequestTemplateDocID) // Gets ID of Request Template doc
    var targetFolder = DriveApp.getFolderById(pendingReponsesFolderID) // Gets ID of Pending Responses folder
    var doc = template.makeCopy(getItemResponse("Last Name", itemResponses) + ", " + getItemResponse("First Name", itemResponses) + " - " + 
      getItemResponse("Title", itemResponses) + ", " + startDate + " to " + endDate, targetFolder)
      
    docID = doc.getId()
    var docVersion = DocumentApp.openById(docID)
    var body = docVersion.getBody()
    var text = body.editAsText()
      
    // Example of finished tableArray: [["Lower School", "Upper School"], "Admissions", "Dean of Faculty"]
    var tableArray = [] 
      
      // If there are files the user submitted, files = true, and vice versa
    var files = false
      
    // Loops through all of the item responses 
    for (var j = 0; j < itemResponses.length; j++) 
    {
      var itemResponse = itemResponses[j]
        
      // Links need to be manually hyperlinked using helper function
      if (itemResponse.getItem().getTitle() == attachFilesSeachTerm) {
        addLinks(itemResponse, body)
        files = true
      }
        
      // Manually codes in case if Department = Lower School (see note above for more details)
      else if (itemResponse.getItem().getTitle() == "Department" && itemResponse.getResponse() == "Lower School") {
        break
      }
        
      // In all other cases, replaces the bracketed text with the form response, be careful of parenthesis (note that this replaces ALL occurences of the bracketed text)
      else if (itemResponse.getItem().getTitle() == "Department" || (itemResponse.getItem().getTitle() == "Division")) {
         tableArray.push(itemResponse.getResponse())
         body.replaceText("{{" + itemResponse.getItem().getTitle() + "}}", itemResponse.getResponse())
       }
        
       else
         body.replaceText("{{" + itemResponse.getItem().getTitle() + "}}", itemResponse.getResponse())     
    }
    
    // If no files are in the item responses, replaces the text with that message
    if (!files)
      body.replaceText("{{" + attachFilesSeachTerm + "}}", "No Files Submitted")
      
    // Special case, Dean of Faculty always needs to sign off on the form so always is added to the array
    tableArray.push("Dean of Faculty")
      
    Logger.log(tableArray)
      
    // Add tables for each person to the doc
    addTables(tableArray, docID, body)
      
    // Share the document with the approvers
    share(tableArray, doc, itemResponses)
  }
  catch (err)
  {
    sendErrorEmail(err, err.lineNumber, docID)
  }
}


// Helper function for formatting title of document, returns the form response to a specific question
function getItemResponse(name, itemResponses)
{
  // Loops through each item response
  // If the item response name matches the name of the specific question, returns the response
  for (var j = 0; j < itemResponses.length; j++) {
    var itemResponse = itemResponses[j];
    if (itemResponse.getItem().getTitle() == name)
        return itemResponse.getResponse()
  } 
  return false
}

// Helper function for adding links to the document
function addLinks(r, b) 
{
  // List of drive file IDs 
  var ids = r.getResponse() 
  
  // Search term
  var search = "{{" + attachFilesSeachTerm + "}}" 
  
  // Finds the element that the search term can be found in. In this case it's the paragraph (every newline is a new paragraph since we use /n)
  var element = b.findText(search)
  
  // Finds the index the search term starts, used as an initial starting position
  var start = element.getStartOffset() 
  
  // Loops through the array of responses. For each response, individually inserts the title of the file as text and hyperlinks the text. This shifts the search term back. Finds the new index of the search term and sets the start position accordingly
  for(var i = 0; i < ids.length; i++) 
  {
    // Gets the name and the url of the file 
    var name = DriveApp.getFileById(ids[i]).getName() + ", "
    var link = DriveApp.getFileById(ids[i]).getUrl()
    
    // Inserts the name of the file at the start position; as a result shifts back the search term
    element.getElement().asText().insertText(start,name)
    
    // Hyperlinks the name (name.length - 3 because name technically includes the space and the comma)
    element.getElement().setLinkUrl(start, start + name.length - 3, link)
    
    // Find the new start position (the element that the search term can be found in)
    element = b.findText(search)
    
    // Set start to new start position
    start = element.getStartOffset()
  }
  
  // Deletes search term (start - 2 to get rid of the extra comma)
  element.getElement().asText().deleteText(start - 2, start + search.length - 1) 
}

// Helper function for getting the value in the cell to the right
// First used to find the name of the Division Director or Department Head (or Dean of Faculty), then used to find that person's email
function getNeighboringCell(cell)
{
  // Gets the "Faculty Table"  spreadsheet
  var url = facultyTableSheetURL
  var sheet = SpreadsheetApp.openByUrl(url)
  
  // Looks for a cell containing a specific text and finds its row and column
  var textFinder = sheet.createTextFinder(cell)
  var row = textFinder.findNext().getRow()
  textFinder = sheet.createTextFinder(cell)
  
  // Adds 1 to the column so you get the neighboring cell
  var column = textFinder.findNext().getColumn() + 1
  
  // Gets the neighboring cell next to the text you found and returns the value of that cell
  var neighbor = sheet.getActiveSheet().getRange(row, column)
  var recipient = neighbor.getValue()
  return recipient;
}

// Helper function for copying the template table and adding it to the end of the doc
function appendTemplateTable(destID)
{
  // Gets the "Comments Box - Template" Google document
  var sourceID = commentsBoxTemplateDocID
  var sourceDoc = DocumentApp.openById(sourceID)
  var sourceBody = sourceDoc.getBody()
  
  // Gets all of the tables in the document and copies the first one
  var tables = sourceBody.getTables()
  var table = tables[0].copy()
  
  // Opens the destination document
  var destDoc = DocumentApp.openById(destID)
  var destBody = destDoc.getBody()
  
  // Appends the template table to the end of the document
  destBody.appendTable(table)
}

// Helper function for adding all of the tables needed
function addTables(responseArray, id, body)
{
  // Loops through the items to search for in the Faculty Table
  for (var i = 0; i < responseArray.length; i++) 
  {
    // If the item is an array (or not a string), loops through the array and for each value, adds a table with that value
    if (typeof responseArray[i] != "string") {
      for (var j = 0; j < responseArray[i].length; j++) {
        actuallyAddTable(responseArray[i][j], id, body)
      }
    } 
    
    // If the item is a string, adds a table with that value
    else
      actuallyAddTable(responseArray[i], id, body)
  }
}

// Helper function for adding a table with the name of a specific Department Head or Division Driector
function actuallyAddTable(departmentOrDivision, docID, body)
{
  // Change this keyword value depending on the word in the "Comments Box - Template" Google doc
  var keyword = keywordSearchTerm 
  
  // Adds the table to the document, and replaces the keyword with the person's name
  appendTemplateTable(docID)
  body.replaceText("{{" + keyword + "}}", getNeighboringCell(departmentOrDivision))
}

// Helper function for sharing a Google doc with the necessary people who will sign off on the request
function share(responseArray, document, itemResponses)
{
  // Array of the emails of people who need to be shared with the document
  var emails = []
  
  // Loops through each value that needs to be found in the table
  for (i = 0; i < responseArray.length; i++)
  {
    // If the item is an array (or not a string), loops through the array and for each value, pushing the corresponding email into the emails array
    if (typeof responseArray[i] != "string") 
    {
      for (var j = 0; j < responseArray[i].length; j++) {
        var name = getNeighboringCell(responseArray[i][j])
        var email = getNeighboringCell(name)
        emails.push(email)
      }
    }
    
    // If the item is a string, adds the corresponding email to the email array
    else 
    {
      var name = getNeighboringCell(responseArray[i])
      var email = getNeighboringCell(name)
      emails.push(email)
    }
  }
  
  // Adds each editor and sends them a notification email with the link to the Google Doc
  var attachmentIds = getItemResponse(attachFilesSeachTerm, itemResponses)
  var docId = document.getId();
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i]
    addEditorSilent(docId, email)
    sendNotificationEmail(email, document.getUrl())
    
    // Makes the approvers editors of all of the attached files
    if (attachmentIds != false) {
      for (var j = 0; j < attachmentIds.length; j++) {
        var attachmentId = attachmentIds[j]
        addEditorSilent(attachmentId, email) 
      }
    }
    
  }
}

// Helper function for adding a specific editor to a Google file without sending that person Google's normal automated email (so Pingry can send its own email)
function addEditorSilent(fileId, userEmail) 
{
  var permissionResource = {
    role: 'writer',
    type: 'user',
    value: userEmail
  };
  
  // Setting this argument to false prevents Google from sending the normal automated email
  var optionalArgs = {
    sendNotificationEmails: false
  };
  
  // Required us to get Advanced Google Services - we got this by going to Resources, then Advanced Google Services, then clicking "on" for "Drive API"
  Drive.Permissions.insert(permissionResource, fileId, optionalArgs);
}

// Helper function for sending a notification email to people that need to sign off on the request that can be customized by Pingry
function sendNotificationEmail(email, link)
{
  // Here is where someone can change the text of the email sent and format it how they want
  var messageBody = "<p> Hello! </p> \
<p> You are receiving this email because you are being asked to approve a Professional Development Request. The link to the \
Google Document can be found by clicking <a href='" + link + "'>here</a>. Please fill out your approval in the \
appropriate box within the next three days. </p> \
<p> Thank you! <\p>"

  // Here is where someone can change details of the message like the recipient, the subject line, and the "name" of the person sending the email
  var message = {
    to: email,
    subject: "Professional Development Request Approval (Action Required)",
    htmlBody: messageBody,
    name: "Automatic Notification Emailer Script"
  };
  
  // Sends the notification email from the account of the person who set the trigger
  MailApp.sendEmail(message);
}

// Helper function for sending an email to a person who submitted the form telling them that the dates the submitted are invalid
function sendInvalidDatesErrorEmail(email)
{
  // Here is where someone can change the text of the email sent and format it how they want
  var messageBody = "<p> Hello! </p> <p> You are recieving this email because your recent Google Form Submission to the form \
'Professional Development Funding Request' contained an error! The dates you submitted were not valid. Please resubmit the form \
with valid dates. </p> <p> Thank you! </p>"

  // Here is where someone can change details of the message like the recipient, the subject line, and the "name" of the person sending the email
  var message = {
    to: email,
    subject: "Your Recent Professional Development Funding Form Submission",
    htmlBody: messageBody,
    name: "Date Validation Emailer Script"
  };
  
  // Sends the notification email from the account of the person who set the trigger
  MailApp.sendEmail(message);
}

// Sends the error email to people if the script malfunctions with the specific line number of the error and a link to the document that might have caused the error
function sendErrorEmail(err, lineNum, docID)
{
  var messageBody = "<p> Oh no! There's been a problem with the Professional Development Form Script. </p>" 
  
  // Adds the line number to the message body
  messageBody += "<p> Line Number: " + lineNum + "</p>"
  
  // Adds a link to the document to the message body
  var url = DocumentApp.openById(docID).getUrl()
  var name = DocumentApp.openById(docID).getName()
  messageBody += "<p> Document: " + "<a href='" + url + "'>" + name + "</a>"
  
  // Adds the computer's error message to the message body
  messageBody += "<p>" + err + "</p>"
  
  // Here is where someone can change details of the message like the recipients, the subject line, and the "name" of the person sending the email
  var message = {
      to: "dshu2019@pingry.org, sburak2020@pingry.org", 
      subject: "Professional Development Form Script Error",
      htmlBody: messageBody,
      name: "Automatic Error Emailer Script"
  }
  
  // Sends the error email from the account of the person who set the trigger
  MailApp.sendEmail(message);
}
