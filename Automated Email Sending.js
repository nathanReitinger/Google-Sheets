// if emails are entered into cells, this script allows
// messages to be sent to those recipients. Keep track
// of Google's email limits per day! 

function Emails() {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName("sheet4");
    var emails = sheet.getRange("D4:D").getValues();
    var Message = sheet.getRange("E4:E").getValues();
    var EmailRArray = [];
    var MessageRArray = [];

for(var i = 0; i < emails.length; i++)    
   if(emails[i][0] != "")
      EmailRArray[i] = emails[i][0];
      
for(var i = 0; i < Message.length; i++)    
   if(Message[i][0] != "")
      MessageRArray[i] = Message[i][0];
     
for(var i = 0; i < EmailRArray.length; i++)
   {
      MailApp.sendEmail(EmailRArray[i], "Message goes here!" + "[optional cell information to put in message" + MessageRArray[i] + "]" + ".");
   }
}

