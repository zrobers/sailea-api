/*
This function is a completely automated version of the SAILea registration process, designed to reduce the workload on the SAILea team, decrease our response time, and increase the number of interactions with new members. The function first collects all of the responses from the last day. It then parses these responses, collecting any useful information for the next steps. Subsequently, the SAILea resources folder is shared with all club leaders and advisors and all emails are added to the SAILea email spreadsheet. Next, a personalized email (using Google Gemini) is sent to the new member. The email invites the new member to SAILea, directs them to our Discord server, and encourages them to set up a meeting with the SAILea team via Zach's Calendly link.
*/
function processResponses() {
   var responses = getRecentResponses();

   for( i = 0; i < responses.length; i++){
    var itemResponses = responses[i];

    var additionalEmails = [];
    var emails = [];

    var indivEmail = itemResponses[1];
    emails.push(indivEmail);
    var firstName = itemResponses[2];
    var lastName = itemResponses[3];
    var schoolName = itemResponses[4];
    var interest = itemResponses[6];
    var hearUs = itemResponses[7];
    var contact = itemResponses[8];

    // share resources folder
    shareFolder(emails);

    // add emails to the spreadsheet
    addEmails(emails);

    // generated LLM-customized part of registration email
    var custom = llmCustomization(firstName, lastName, schoolName, interest);
    Logger.log(custom);

    // send the registration email
    additionalEmails.push('zach.robers@gmail.com');
    additionalEmails.push('aritras059@gmail.com');
    additionalEmails.push('hanqixiao.personal@gmail.com');
    additionalEmails.push('tony.nunn@duke.edu');
    sendRegistrationEmail(firstName, lastName,indivEmail, additionalEmails, custom);

   }
   
}

/*
Returns all of the responses on the SAILea Club Registration Form within the last day 
*/
function getRecentResponses() {
  var sheetId = '1AmFQjaSDTe6rMvbPqXJjZI6XZBVuyDj94HKRUx2bhZA';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName('Form Responses 1');
  var lastRow = sheet.getLastRow();
  var currentDate = new Date();
  var twentyFourHoursAgo = new Date(currentDate.getTime() - 24 * 60 * 60 * 1000);

  var rowsWithin24Hours = [];

  // Iterate through each row starting from the second row
  for (var i = 2; i <= lastRow; i++) { // Start from 2 to skip header row
    var cell = sheet.getRange(i, 1).getValue();
    
    // Check if the cell is not empty, is a date, and is within the past 24 hours
    if (cell && cell instanceof Date && cell >= twentyFourHoursAgo && cell <= currentDate) {
      rowsWithin24Hours.push(sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0]); // Collect entire row
    }
  }

  return rowsWithin24Hours;
}

/*
Sends the registration email to the club leader with any additional advisors and club leaders CC'd
*/
function sendRegistrationEmail(first, last, primaryEmail, toCC, customizedSection){
  var fullName = first + ' ' + last
  var recipient = primaryEmail.toString();
  Logger.log(recipient);
  var subject = "Welcome to SAILea";
  var body = `<p>Dear ${fullName},<p>

<p>My name is Zach Robers and I am an undergrad at Duke, and the Chairman of the Scholastic Artificial Intelligence League. I would like to personally welcome you to the SAILea community. We are so excited that you could join us!<p>

<p>${customizedSection}<p>

<p>As you likely know, SAILea offers AI resources targeted at a high school-level audience. I have shared our Google Drive folder containing these resources with your email and any additional emails you provided on the form. We've expanded our resource offerings throughout the school year to include exciting new interactive activities on topics like the Stock Market, the ChatGPT API, and the creation of your own Custom GPT. If there is anything you would like for us to make a resource on, feel free to let me know.<p>

<p>SAILea also offers courses in programming and AI. We have created 4 courses in Python, Java, The Mathematics Behind Deep Learning, and The Principles of Machine Learning and Deep Learning. Periodically we host live lessons on various topics. Learn more at https://www.sailea.org/courses. Course recordings and resources can be found in the resource folder that has been shared with you.<p>

<p>In addition to offering resources and courses, SAILea hosts speaker events and competitions. Our last speaker event featured UNC-Chapel Hill Professor Richard Marks on combining VR and generative AI. The recording is available on YouTube to view. We are also planning additional speaker events and competitions for later this year. We would love it if you could join us for these events and feel free to invite anyone else as well. More details to come soon. Also, please join our Discord server (https://discord.gg/y22aTa4a2f) where we will post important announcements and details about upcoming events.<p>

<p>Good luck with your AI/CS journey, and I am looking forward to hearing from you.<p>

<p>Zach<br>
Chair<br>
The Scholastic Artificial Intelligence League<p>
`;
  var ccAddresses = toCC.join(',');

  // Send the email with CC addresses
  GmailApp.sendEmail(recipient, subject, '', {
    cc: ccAddresses,
    htmlBody: body
  });
}

/*
Shares the SAILea resources folder with the club leader, advisor, and additional leaders
*/
function shareFolder(emails) {
  var folder = DriveApp.getFolderById('1JYMUgXSd_qQy2QxMKSXsynxAJ39CgKJC');
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i];
    try {
      // Add user with specified permission
      folder.addViewer(email); 
    } catch (e) {
      // Log any errors encountered
      Logger.log("Error sharing with " + email + ": " + e);
    }
  }
}

/*
Generates a personalized section of the email addressing any noteworhty aspects of their club using Gemini LLM
*/
function llmCustomization(first, last, schoolName, interest, ) {
  var fullName = first + ' ' + last
  var q = `Your task is to act on behalf of SAILea, a nonprofit dedicated to helping students interested in artificial intelligence (AI) through AI resources such as presentations on NLP, machine learning, computer vision, etc., courses in Python, Java, and ML, speaker events, and advising. Write a personalized 3 sentence paragraph with a friendly tone to ${fullName}, a student at ${schoolName}, commenting on their specific interest: ${interest}. Do not include a greeting. This text will be sent to the student in a welcome email to SAILea. Write three sentences to the student in a single paragraph following these explicit instructions.`;

  if(schoolName == 'N/A' || schoolName == ''){
    q = `Your task is to act on behalf of SAILea, a nonprofit dedicated to helping students interested in artificial intelligence (AI) through AI resources such as presentations on NLP, machine learning, computer vision, etc., courses in Python, Java, and ML, speaker events, and advising. Write a personalized 3 sentence paragraph with a friendly tone to ${fullName}, commenting on their specific interest: ${interest}. Do not include a greeting. This text will be sent to the student in a welcome email to SAILea. Write three sentences in a single paragraph following these explicit instructions.`;
  }

  const apiKey = "AIzaSyChDRHh-88-n27zfZBCHoVYp8VpsdcnznU"; // Please set your API key.

  const url = `https://generativelanguage.googleapis.com/v1/models/gemini-pro:generateContent?key=${apiKey}`;
  const payload = { contents: [{ parts: [{ text: q }] }] };
  const res = UrlFetchApp.fetch(url, {
    payload: JSON.stringify(payload),
    contentType: "application/json",
  });
  const obj = JSON.parse(res.getContentText());
  if (
    obj.candidates &&
    obj.candidates.length > 0 &&
    obj.candidates[0].content.parts.length > 0
  ) {
    return obj.candidates[0].content.parts[0].text;
  } else {
    return "";
  }
}


/*
Adds all emails to the SAILea email database
*/
function addEmails(emails){
  var sheet = SpreadsheetApp.openById('1d2uhM0tf5P-W_93vo5rO0RU-84cQ73ItJ9aekDsRjn4').getActiveSheet();
  var columnA = sheet.getRange("A:A").getValues(); 
  var lastRow = columnA.filter(String).length + 1; 
    for (var i = 0; i < emails.length; i++) {
      sheet.getRange(lastRow + i, 1).setValue(emails[i]); 
    }
}

