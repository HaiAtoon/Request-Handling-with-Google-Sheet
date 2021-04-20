function getData(value) {


      var request = {};


      request.studentName = value[2];
      request.id = value[3];
      request.requestType = value[7];
      request.requestStatus = value[11];
      request.requestComment = value[12];
      
      if(value[12]=='')
        request.requestComment='אין הערות';

  return request;
}

function toHtml(data){

  var htmlTemplate = HtmlService.createTemplateFromFile("template.html");
  htmlTemplate.request = data;
  var htmlBody = htmlTemplate.evaluate().getContent();
  
  return htmlBody;

}

function sendEmailNew(){

  var startRow = 2;
  var numRows = 1000;
  var subject = 'הודעה בנוגע לבקשה ששלחת';
  var count=0;
  var row=2;


  var values = SpreadsheetApp.getActive().getSheetByName("בקשות חדשות").getRange(startRow, 1, numRows, 14).getValues();

  values.forEach(function(value) {

    
    if((value[11]=='מאושר'||value[11]=='נדחה')&&value[13]!='נשלחה הודעה'){
      var emailAddress = value[4];

      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: toHtml(getData(value))
        });

      SpreadsheetApp.getActiveSheet().getRange('N'+row).setValue('נשלחה הודעה');
      SpreadsheetApp.getActiveSheet().getRange(row,1,1,14).setBackgroundRGB(169,169,169);

      count++;
      
    }

    row++;

  });


  if(count == 0)
    Browser.msgBox("אין הודעות חדשות לשליחה");
  else
    Browser.msgBox("נשלחו "+count+" הודעות");
}
