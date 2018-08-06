function AutofillRecap() {
  var day = 1000 * 60 * 60 * 24;
  var ss = SpreadsheetApp.openByUrl(
   'https://docs.google.com/spreadsheets/d/17qD_2qCHmDWW_jhPOouuCmdykFSq2w4zN-1hoFv1Jho/edit?usp=sharing').getActiveSheet();
  var strategySheet = SpreadsheetApp.openByUrl(
   'https://docs.google.com/spreadsheets/d/1ZMlmfgZkmSyXwfjmacj6cm9WmW35JIZ4jm1Wp9kMMV0/edit?usp=sharing').getActiveSheet();
  
  function Student(name, username, date, ixl, time, nextSession, tracker) {
    this.name = name;
    this.username= username;
    this.date = date;
    this.time = time;
    this.ixl = ixl;
    this.nextSession = nextSession;
    this.tracker = tracker;
  }; 
 
 var test = new Student ("Test", "testburroughs417", ss.getRange(8,3).getValue(), ss.getRange(8,4).getValue(),ss.getRange(8,5).getValue(), "with Coach Chris @ TAC's Offices",
                         "https://docs.google.com/spreadsheets/d/1kRIzF0EQfYtERx7o1G6KOq0n03p0lSEPD5lnQoumvP4/edit?usp=sharing");
 var gabe = new Student ("Gabe","gabeherschelman417", ss.getRange(4,3).getValue(), ss.getRange(4,4).getValue(), ss.getRange(4,5).getValue(), "with Coach Chris @ TAC's Offices", 
                          "https://docs.google.com/spreadsheets/d/1GMl3NDXILQ5tG5tKwkjyxuTObZMyNH9-RFKjcz_xvvc/edit?usp=sharing");
 var joey = new Student ("Joey", "joeyburroughs417", ss.getRange(8,3).getValue(), ss.getRange(8,4).getValue(),ss.getRange(8,5).getValue(), "with Coach Chris @ TAC's Offices",
                         "https://docs.google.com/spreadsheets/d/1R_zSUoBOe5lbaSppOuXx7_GZ_mfK7kbHf9boS-N4qy0/edit?usp=sharing");
 var emily = new Student ("Emily", "emilyburroughs417", ss.getRange(10,3).getValue(), ss.getRange(10,4).getValue(),ss.getRange(10,5).getValue(), "with Coach Chris @ TAC's Offices", 
                         "https://docs.google.com/spreadsheets/d/13-xSbSBNEHNEvewxMO5IvMziYE_7-PnrJDJQDPArhlw/edit?usp=sharing");
  
  var assignment = {
    '{ENGLISH}':ss.getRange(2,1).getValue(), 
    '{MATH}':ss.getRange(2,2).getValue(), 
    '{READING}':ss.getRange(2,3).getValue(),
    '{SCIENCE}':ss.getRange(2,4).getValue(),
    '{ENGLISH AND MATH}':ss.getRange(2,5).getValue(),
    '{READING AND SCIENCE}':ss.getRange(2,6).getValue(),
    '{FULL TEST}':ss.getRange(2,7).getValue(),
    '{ENGLISH REVIEW}':strategySheet.getRange(2,1).getValue(),
    '{MATH REVIEW}':strategySheet.getRange(2,2).getValue(),
    '{READING REVIEW}':strategySheet.getRange(2,3).getValue(),
    '{SCIENCE REVIEW}':strategySheet.getRange(2,4).getValue(),
                   }
  
  var strategyLinks = {
    "ACT English Rules to Remember: STOP":"https://drive.google.com/open?id=1ojKzg3Qg9KjXsHS-Jw9PVQv72yc1f3dnfGqjUnWKyfs",
    "ACT Math to Memorize: SUPERB":"https://drive.google.com/open?id=1ONzrEqVWQqYzlhacKQr3mi3eVglg6cCOrILKBMrDfLA",
    "ACT Math Cheat Sheet":"https://drive.google.com/open?id=1T-NW1jjjgqsdAbusUpfqHuUZQctcP6YvhMX5sPFXdO8",
    "ACT Reading: Be SMART":"https://drive.google.com/open?id=1uN9QuqowyeF6EXNOJM77vS7fpEP3I1Nvd2Tic6yVBqc",
    "ACT Science: SLAP: 32+":"https://drive.google.com/open?id=1dUeDA8fn5GGvdz7vK0376b4eXiye7xuKqOjkkWa3JZo",
  }

  function replaceText(student) {
    var doc = DocumentApp.getActiveDocument().getBody(); 
    var trackerSheet = SpreadsheetApp.openByUrl(student.tracker).getActiveSheet();
    for(var section in assignment) {
      var search = doc.findText(section) 
      if(search != null){
        doc.replaceText(section, assignment[section])
        for (var col=1;col<12;col++) {
          var columnCounter = trackerSheet.getRange(1,col);
          if (columnCounter.getValue() == section) {
          var rowCounter = "rowCounter";
          for (var row=1; row<100;row++) {
            if(rowCounter != null) {
              rowCounter = trackerSheet.getRange(row,col);
              if (rowCounter.getValue() == 2) {
                rowCounter.setBackground("red");
                rowCounter.setValue(1);
                rowCounter = null;
              }
            }
          }
        }
      }
    }
  }
    
  function linkText (unlinkedText) {
	if(doc.findText(unlinkedText) != null) {
	  var textElement = doc.findText(unlinkedText).getElement();
      var textStart = textElement.asText().findText(unlinkedText).getStartOffset();
      var textEnd = textElement.asText().findText(unlinkedText).getEndOffsetInclusive();
      textElement.setLinkUrl(textStart, textEnd, strategyLinks[unlinkedText]);	
      Logger.log(strategyLinks[unlinkedText]);
	}
  }
    linkText("ACT English Rules to Remember: STOP");
    linkText("ACT Math to Memorize: SUPERB");
    linkText("ACT Math Cheat Sheet");
    linkText("ACT Reading: Be SMART");
    linkText("ACT Science: SLAP: 32+");
    
    function correctFormat(date) {
      return Utilities.formatDate((date), "GMT", "E M/d");
    }
    var firstDate = new Date(student.date);
    var datePlus7 = correctFormat(new Date(firstDate.getTime() + 7 * day));
    var datePlus14 = correctFormat(new Date(firstDate.getTime() + 14 * day));
    var datePlus21 = correctFormat(new Date(firstDate.getTime() + 21 * day));
    firstDate = correctFormat(firstDate);
	
	//firstDate has to be formatted last b/c other dates use firstDate as a date object; formatting converts to string.
	//While I could use a loop to put as many +7 dates as I want into an array and then print each element, I never need to print
	//more or fewer than four dates, so the added complexity would be unnecessary.
	
    doc.replaceText('{DATE}', firstDate + " "  + student.time + " " + student.nextSession + 
                    "\n" + datePlus7 + " " + student.time + " " + student.nextSession + "\n" + datePlus14 + " " + 
                    student.time + " " + student.nextSession + "\n" + datePlus21 + " " + student.time + " " + student.nextSession);
    doc.replaceText('{NAME}', student.name);  
    doc.replaceText('{USERNAME}', student.username);
    doc.replaceText('{IXL}', student.ixl);
  }
 replaceText(test); 
}