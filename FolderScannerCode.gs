/*
Folder Scanner Script

Version 1.0
Dillon Shu, Samantha Burak
8/1/19

Trigger runs every day at around 6 am (will be changed to run Mondays, Wednesdays, and Fridays)

Description:
Scans the pending folder to see if the documents in it have been completed (all of the comment boxes and signature boxes have been filled out
If a document is completely filled out, moves it to the completed folder and sends an email to the Head of School
If not, and it's been three days since the form was submitted, sends reminder emails to the people that still need to fill the document out
If a document has been approved by the Head of School, moves it to the head of school approved responses folder and sends an email to the Dean of Faculty
*/

// Global vars
var pendingFolderID = '1FQajZxxrfNM6Znt_X49da4qlgVeEHNiY'
var completedFolderID = '1Kw6QalRj0zjcXKikpbxGbdDzuXolndBh'
var headOfSchoolApprovedFolderID = '1S4hXiUP5sgsyprl98Z--uuUT7VEp_MDR'
var facultyTableSheetURL = 'https://docs.google.com/spreads/d/127XQ0elfDQiyH1wjgoltNilRROAO3U_c2bSqVoJS7m4/edit#gid=0'

// Keep track of the person that will be sent a reminder email for the purpose of error message handling 
var potentialPersonName = ""

// Main function of this program
function checkFolder() {
  var docID = ""
  try 
  {
    // Get the folders for completed requests and pending requests
    var pendingFolder = DriveApp.getFolderById(pendingFolderID)
    var completedFolder = DriveApp.getFolderById(completedFolderID)
    var headOfSchoolApprovedFolder = DriveApp.getFolderById(headOfSchoolApprovedFolderID)
    
    // Gets all files in the pendingFolder, returns type FileIterator
    var files = pendingFolder.getFiles()
    
    // Gets the spreadsheet "Faculty Table"
    var url = facultyTableSheetURL
    var sheet = SpreadsheetApp.openByUrl(url)
    
    // Creates a list of the names of all of the approvers/faculty members from the spreadsheet
    // Do NOT change the value in firstCell. Value in lastCell changes automically if someone adds a row to the spreadsheet
    var firstCell = "B3"
    var lastCell = "B" + sheet.getLastRow()
    var listOfNames = sheet.getRange(firstCell + ":" + lastCell).getValues()
    listOfNames = cleanUpListOfNames(listOfNames) 
    
    // Creates a dictionary with all of the keys being each faculty member's name and sets all the values to empty arrays
    var map = {}
    for (var i = 0; i < listOfNames.length; i++)
    {
      if (listOfNames[i].length > 0)
        map[listOfNames[i]] = []
    }
    
    // For each document, checks whether or not it has been completed
    while (files.hasNext()) 
    {
      var file = files.next()
      docID = file.getId()
      checkApproval(pendingFolder, completedFolder, docID, map)
    }
   
    // Sends reminders to everyone who still needs to comment on a doc
    sendReminders(map, listOfNames)
    
    var completedDocs = completedFolder.getFiles()
    
    while (completedDocs.hasNext())
    {
      var completedDoc = completedDocs.next()
      docID = completedDoc.getId()
      checkHeadApproval(completedFolder, headOfSchoolApprovedFolder, docID)
    }
      
   
  }
  catch (err)
  {
    sendErrorEmail(err, err.lineNumber, docID, potentialPersonName)
  }
}

// Sends all of the reminders to everyone who still needs to common on a doc
function sendReminders(map, listOfNames)
{
  for (var i = 0; i < listOfNames.length; i++)
  {
    var docIDs = map[listOfNames[i]]
    if (docIDs.length > 0) {
      potentialPersonName = listOfNames[i]
      sendReminderEmail(getNeighboringCell(listOfNames[i]), docIDs)
    }
  }
}

// Checks to see if Head of School has signed his/her name on a Completed Response file. If yes, moves the doc to the corresponding folder and
// sends an email to the Dean of Faculty
function checkHeadApproval(completedFold, headApprovedFold, docID)
{
  var doc = DocumentApp.openById(docID)
  var header = doc.getHeader()
  if (header != null)
  {
    var headerText = header.getText()
    // Checks to see if header includes "Head of School" - should it check for the words "Approved" and "Denied"?
    if (headerText.indexOf('Head of School') >= 0){
      Logger.log(headerText)
      moveDocToNewFolder(docID, completedFold, headApprovedFold)
      sendEmailToDeanOfFaculty(doc.getUrl(), doc.getName())
    }
  }
}

// Gets rid of the extra spaces and "x"s in a list
function cleanUpListOfNames(list)
{
  newList = []
  for (var i = 0; i < list.length; i++)
  {
    var name = list[i]
    if (name != "" && name != "x")
      newList.push(name)
  }
  return newList
}

// Adds a document ID to the dictionary
function addToDict(dict, docID, name)
{
  var arr = dict[name]
  arr.push(docID)
  dict[name] = arr
}

// Checks if a document has been signed off by the necessary faculty members
function checkApproval(pendingFolder, completedFolder, currentDocID, dict) {
  
  // Gets all of the tables that are currently in the doc
  var sourceDoc = DocumentApp.openById(currentDocID)
  var sourceBody = sourceDoc.getBody()
  var tables = sourceBody.getTables()
  
  Logger.log(DocumentApp.openById(currentDocID).getName())
  
  // Will contain the indicies of the tables that correspond to people who have yet to comment on the doc
  var shame = [] 
  
  // Loops through every table in the document (except for the first one which is always the form response)
  for (var i = 1; i < tables.length; i++) 
  {
    // Gets the cell next to the one with "Signature" - right now that's (2, 1) - and removes all spaces that might interfere with counting the length
    var cellText = tables[i].getCell(tables[i].getNumRows() - 1, 1).getText()     
    cellText = cellText.replace(/\s+/g, '')
      
    // If there isn't text in the cell, adds the table index to the shame array
    if (cellText.length == 0) 
     shame.push(i)   
  }
 
  // If there is no one who hasn't signed off on the request, move the document to the completedFolder and send an email to the Head of School
  if (shame.length == 0) {
    moveDocToNewFolder(currentDocID, pendingFolder, completedFolder)
    sendCompletedDocEmail(sourceDoc.getUrl(), sourceDoc.getName())
  }
  else
  {
    // Checks if the document is more than 3 days old
    if (getDiffBetweenDates(currentDocID) >= 3.0)
    {
      // Registers all of the people that still needs to comment on the doc in the dictionary
      for(var i = 0; i < shame.length; i++)
      {
        var table = tables[shame[i]]
        // Hardcoded to get the text in the first row of the table (technically (0, 0))
        var nameOfPerson = table.getCell(0, 0).getText()
        addToDict(dict, currentDocID, nameOfPerson)
      }
    }
  }
}

// Returns the difference between the current date and the date a document was created
function getDiffBetweenDates(docID)
{
  // Gets both dates
  var file = DriveApp.getFileById(docID)
  var dateCreated = file.getDateCreated()
  var today = new Date()
  
  // Gets both dates in milliseconds since 1970
  var dateCreatedTime = dateCreated.getTime()
  var todayTime = today.getTime()
  
  // 24 * 3600 * 1000 is number of milliseconds in a day
  var diffInDays = (todayTime - dateCreatedTime)/(24 * 3600 * 1000)
  
  return diffInDays
}

// Move a doc from one folder to a different folder
function moveDocToNewFolder(docID, oldFolder, newFolder)
{
  var file = DriveApp.getFileById(docID)
  newFolder.addFile(file)
  oldFolder.removeFile(file)
}

// Gets a cell to the right of the given cell (taken from the Pro Development Form Script, with a few changes)
function getNeighboringCell(cell)
{
  // Gets the "Faculty Table" spreadsheet
  var url = facultyTableSheetURL
  var sheet = SpreadsheetApp.openByUrl(url)
  
  // Looks for a cell containing a specific text and finds its row and column
  var textFinder = sheet.createTextFinder(cell);
  var next = textFinder.findNext()
  var row = next.getRow()
  
  // Adds 1 to the column so you get the neighboring cell
  var column = next.getColumn() + 1
  
  // Gets the neighboring cell next to the text you found and returns it
  var neighbor = sheet.getActiveSheet().getRange(row, column)
  var recipient = neighbor.getValue()
  return recipient
}

// Makes sure that the IDs of the files are sorted in ascending order
function getArrayOfDates(arrayOfIDs)
{
  var arrayOfDates = []
  for (var i = 0; i < arrayOfIDs.length; i++)
  {
    var file = DriveApp.getFileById(arrayOfIDs[i])
    var dateCreated = file.getDateCreated()
    arrayOfDates.push(dateCreated)
  }
  return arrayOfDates
}

// Make the IDs of the files sorted by dates created
function getSortedIDs(arrayOfIDs)
{
  var sortedIDs = []
  
  // Creates a dictionary with keys as string versions of dates created and values as the corresponding file ID
  var fileDict = {}
  for (var i = 0; i < arrayOfIDs.length; i++)
  {
    var ID = arrayOfIDs[i]
    var file = DriveApp.getFileById(ID)
    var dateCreated = file.getDateCreated()
    fileDict[dateCreated] = ID
  }
  
  // Creates an array with the keys of the dictionary (string versions of dates created)
  var stringDates = Object.keys(fileDict)
  
  // Creates a new array that contains the actual Date objects with dates created
  var dates = []
  for (var i = 0; i < stringDates.length; i++)
    dates.push(new Date(stringDates[i]))
    
  // This is a comparison function that will result in dates being sorted in ascending order
  var date_sort_asc = function (date1, date2) {
    if (date1 > date2) return 1
    if (date1 < date2) return -1
    return 0
  }
  
  // Sort the dates in ascending order
  dates.sort(date_sort_asc)
  
  // Gets the IDs of those files in that sorted order and pushes them to the sortedIDs array
  for (var i = 0; i < dates.length; i++)
  {
    var sortedDate = dates[i]
    sortedIDs.push(fileDict[sortedDate])
  }
  
  return sortedIDs
}

// Sends the reminder email to people who have not yet signed off on the approval docoument
function sendReminderEmail(email, arrayOfIDs)
{
  // Make the IDs of the files sorted by dates created
  var sortedIDs = getSortedIDs(arrayOfIDs)
  
  Logger.log("--------------------")
  Logger.log("Email: " + email)
  Logger.log("dates should be sorted: " + getArrayOfDates(sortedIDs))
 
  // Creates a bulletpoint list of all of the documents that person still needs to comment on
  var bulletPoints = "<ul> "
  for (var i = 0; i < sortedIDs.length; i++)
  {
    var ID = sortedIDs[i]
    var url = DocumentApp.openById(ID).getUrl()
    var name = DocumentApp.openById(ID).getName()
    bulletPoints += "<li> <a href='" + url + "'>" + name + "</a> </li>"
  }
  bulletPoints += " </ul>"
  
  // Here is where someone can change the text of the email sent and format it how they want
  var messageBody = "<p> Hello! </p> \
<p> You are receiving this email as a reminder that you need to comment on one or more Professional Development Requests. \
The Google Document(s) you need to go to is/are: </p> <p> " + bulletPoints + " </p> <p> Please fill out your comment(s) and signature(s) in the appropriate box(es) ASAP. </p> \
<p> Thank you! <\p>"

  // Here is where someone can change details of the message like the recipient, the subject line, and the "name" of the person sending the email
  var message = {
    to: email,
    subject: "Reminder for Professional Development Request Approval (Action Required)",
    htmlBody: messageBody,
    name: "Automatic Reminder Emailer Script"
  }
  
  // Send the reminder email from the account of the person who set the trigger
  MailApp.sendEmail(message);
}

// Sends an email to the Head of School when a document has been moved to the Completed folder
// Added name parameter to include name of doc in email subject. That way, the emails won't thread.
function sendCompletedDocEmail(link, name)
{
  var messageBody = "<p> Hello! </p> \
<p> You are receiving this email because a Professional Development Request has received its necessary signatures. \
The link to the Google Document can be found by clicking <a href='" + link + "'>here</a>. \
<p> Thank you! <\p>"
  
  // Here is where someone can change details of the message like the recipients, the subject line, and the "name" of the person sending the email
  var message = {
      to: "bburkhart@pingry.org",
      subject: "Completed Professional Development Request: " + name,
      htmlBody: messageBody,
      name: "Automatic Completed Pro Grow Request Emailer Script"
  }
  
  // Sends the error email from the account of the person who set the trigger
  MailApp.sendEmail(message)
}

// Sends an email to Dean of Faculty when a document has been approved by the Head of School
function sendEmailToDeanOfFaculty(link, name)
{
  var deanOfFacultyEmail = getNeighboringCell(getNeighboringCell("Dean of Faculty"))
  Logger.log(deanOfFacultyEmail)
      
  var messageBody = "<p> Hello! </p> \
<p> You are receiving this email because a Professional Development Request has been approved by the Head of School. \
The link to the Google Document can be found by clicking <a href='" + link + "'>here</a>. \
<p> Thank you! <\p>"
  
  // Here is where someone can change details of the message like the recipients, the subject line, and the "name" of the person sending the email
  var message = {
      to: deanOfFacultyEmail,
      subject: "Head of School Approved Professional Development Request: " + name,
      htmlBody: messageBody,
      name: "Automatic Head of School Approved Pro Grow Request Emailer Script"
  }
  
  // Sends the error email from the account of the person who set the trigger
  MailApp.sendEmail(message)
}


// Sends the error email to people if the script malfunctions with the specific line number of the error and a link to the document that might have caused the error
function sendErrorEmail(err, lineNum, docID, potentialNameOfPerson)
{
  // Gets the current name of the script
  var scriptID = ScriptApp.getScriptId()
  var scriptName = DriveApp.getFileById(scriptID).getName()

  var messageBody = "<p> Oh no! There's been a problem with your script, " + scriptName + ". </p>" 
  
  // Adds the line number to the message body
  messageBody += "<p> Line Number: " + lineNum + "</p>"
  
  // Adds a link to the document to the message body
  var url = DocumentApp.openById(docID).getUrl()
  var name = DocumentApp.openById(docID).getName()
  messageBody += "<p> If there was an issue with checking a document, that document might have been: " + "<a href='" + url + "'>" + name + "</a> </p>"
  messageBody += "<p> If there was an issue with sending a reminder email to a person, that person might have been: " + potentialNameOfPerson + "</p>"
  
  // Adds the computer's error message to the message body
  messageBody += "<p>" + err + "</p>"
  
  // Here is where someone can change details of the message like the recipients, the subject line, and the "name" of the person sending the email
  var message = {
      to: "sburak2020@pingry.org", 
      subject: scriptName + " Error",
      htmlBody: messageBody,
      name: "Automatic Error Emailer Script"
  }
  
  // Sends the error email from the account of the person who set the trigger
  MailApp.sendEmail(message);
}
