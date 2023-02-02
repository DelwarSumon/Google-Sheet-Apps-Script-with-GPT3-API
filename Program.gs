/**
 * This is the mail function to control this program
 * Generate complete text, retriev emails from sheet, send mail, push to sheet
 */
function main(){
  var prompt = "Write a poem about What is the meaning of life?";
  var generatedText = generateText(prompt);

  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Sheet1');
  var subject = "Test email from Google AppScript";
  var emailArr = retrievEmails();

  var now = new Date();
  var dateTimeString = now.toLocaleString();
  var sheetArr = [];

  var message = "This is Programatically generated Poem.\n Developed By: Delwar Sumon\n______________________________________________\n\n" + generatedText;

  for (var i = 0; i < emailArr.length; i++) {
    //Send email to recipient
    var emailSent = sendEmailToEmail(emailArr[i], subject, message);
    // Prepare data to push to sheet
    var dataArr=[];
    dataArr.push(dateTimeString);
    dataArr.push(emailArr[i]);
    dataArr.push(message);
    dataArr.push((emailSent) ? 'Sent' : 'Not_Send');
    sheetArr.push(dataArr);
  }
  if(sheetArr.length > 0){
    // Push all data to sheet
    sheet.getRange(sheet.getLastRow()+1, 1, sheetArr.length, sheetArr[0].length).setValues(sheetArr);
  }
}

/**
 * Generate a complete text from GPT3 API
 */
function generateText(prompt) {
  var apiKey = "YOUR_API_KEY";
  
  var url = "https://api.openai.com/v1/engines/davinci/completions";
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify({
      "prompt": prompt,
      "max_tokens": 100,
      "n": 1,
      "stop": null,
      "temperature": 0.5,
    })
  };
  
  var response = UrlFetchApp.fetch(url, options);
  return JSON.parse(response.getContentText()).choices[0].text;
  
}

/**
 * Pull emails from Sheet2
 */
function retrievEmails(){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Sheet2');
  // var range = sheet.getRange("A2:A1000");
  // var data = range.getValues();
  var data = spreadsheet.getDataRange().getValues();
  var emailArr = [];
  for (var i = 0; i < data.length; i++) {
    if(validateEmail(data[i])){
      emailArr.push(data[i]);
    }
  }

  return emailArr;
}

/**
 * Validate email
 */
function validateEmail(email) {
  var emailRegex = /^[a-zA-Z0-9._-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,4}$/;
  return emailRegex.test(email);
}

/**
 * Send mail to recipient email
 */
function sendEmailToEmail(recipient, subject, message) {
  try {
    GmailApp.sendEmail(recipient, subject, message);
    return true;
  } catch (e) {
    Logger.log("Error sending email: " + e.message);
    return false;
  }
}




