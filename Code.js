var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
var ui = SpreadsheetApp.getUi()

// VALUES TO EDIT
var CELL_FOR_NAME = 'Project Manager'
var EMAIL = '@gmail.com'
var SUBJECT = "Write the subject of the email here."


function getValues() {
  // GETS INITIAL VALUES FROM CELLS IN ACTIVE ROW
  let row = ss.getActiveRange().getRowIndex()
  
  // NAMES OF CELLS TO READ DATA FROM
  let name = getByName("Project Manager", row);
  let nameArray = nameSplit(name);
  
  let emailArray = getEmails(nameArray);
  let client = getByName("Client Name", row);
  let jobName = getByName("Job Name", row);
  let jobNum = getByName("Job Number", row);

  sendEmail(nameArray, emailArray, client, jobName, jobNum);
}


function nameSplit(name) {
  // CHECKS FOR MULTIPLE PROJECT MANAGERS
  if (!name.includes(',')){
    return splitBySpace(name);
  } else {
    return splitByComma(name);
  }
}


function splitByComma(name) {
  let splitName = name.split(',');
  return splitName
}


function splitBySpace(name) {
    var splitName = name.split(" ");
    return [splitName[0]];
}


function getEmails(nameArray) {
  // CONVERTS NAMES FROM nameArray INTO EMAIL ADDRESSES FOR emailArray
  var emailArray = [];

  for (let i = 0; i < nameArray.length; i++) {
    let email = nameArray[i].concat('', EMAIL);
    emailArray.push(email);
  }
  return emailArray
}

function getByName(colName, row) {
  // GETS NAME CELL DATA BASED ON COLUMN NAME AND ACTIVE ROW
  var data = ss.getDataRange().getValues();
  var col = data[0].indexOf(colName);
  if (col != -1) {
    return data[row-1][col];
  }
}


function sendEmail(names, emails, client, job, num) {
  // GIVE USER A PROMPT
  let response = ui.alert('You are sending an email to:\r\n ' + emails, 
  ui.ButtonSet.OK_CANCEL);

  // READS PROMPT RESPONSE
  if (response == ui.Button.OK) {
    for (let i = 0; i < emails.length; i++) {

      // SENDS EMAIL
      MailApp.sendEmail({
        to: emails[i],
        subject: SUBJECT,
        htmlBody: "Dear " + names[i] + ",<br>Your project #" 
        + num + " " + job + " with " + client + 
        " is complete!<br>If your job is being installed, the installation department has already been notified.",
      })
    }
    SpreadsheetApp.getActive().toast('Email(s) sent.');
  }
