// Updates the Mass capacity options
function massLimiter() {


  // Constants
  var massOcc = 56;
  var numOfSlots = 23;

  // Get curr spreadsheet for active form resp
  var SHEET_ID = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById(SHEET_ID);  //Responses spreadsheet with all masses
  var form = FormApp.getActiveForm();  // Current form
  var sheet = ss.getSheetByName('Master');  // Active Sheet

  var numResp = sheet.getLastRow();
  var timestamp = sheet.getRange(2, 1, numResp).getValues();
  var massAtt = sheet.getRange(2, 6, numResp).getValues();
  var massCount = sheet.getRange(2, 5, numResp).getValues();
  
  var today = new Date();
  var startOfWeek = new Date();
  var days = today.getDay();
  if (days == 0) {
    days = 7;
  }
  startOfWeek.setDate(today.getDate() - (days - 1));
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  Logger.log(days);
  var massArr = [0, 0, 0, 0];
  var numMass = [0, 0, 0, 0];

  var times = ['Saturday 5pm','Sunday 10am','Sunday 5pm','Sunday 9pm'];

  for (var ind = 0; ind < massAtt.length; ind++) {
    var resp = massAtt[ind];
    var currDate = timestamp[ind].toString().split(' ');
    var currMonth = -1;
    for (var j = 0; j < 12; j++) {
      if (currDate[1] === months[j]) {
        currMonth = j;
        break;
      }
    }
    var currDay = new Date(parseInt(currDate[3]), currMonth, parseInt(currDate[2]));
    var isValid = false;
    if (currDay.getFullYear() > startOfWeek.getFullYear()) {
      isValid = true;
    } else if (currDay.getFullYear() == startOfWeek.getFullYear()) {
      if (currDay.getMonth() == startOfWeek.getMonth()) {
        if (currDay.getDate() >= startOfWeek.getDate()) {
          isValid = true;
        }
      } else if (currDay.getMonth() > startOfWeek.getMonth()) {
        isValid = true;
      }
    }
    if (isValid) {
      if (resp.toString().search(times[0]) != -1) {
        massArr[0] += massCount[ind][0];
        numMass[0]++;
      } else if (resp.toString().search(times[1]) != -1) {
        massArr[1] += massCount[ind][0];
        numMass[1]++;
      } else if (resp.toString().search(times[2]) != -1) {
        massArr[2] += massCount[ind][0];
        numMass[2]++;
      } else if (resp.toString().search(times[3]) != -1) {
        massArr[3] += massCount[ind][0];
        numMass[3]++;
      }
    }
  }
  
  Logger.log(massAtt);
  Logger.log(massCount);
  Logger.log(massArr);
  Logger.log(numMass);

  var list = [];
 
  for (var i = 0; i < times.length; i++) {
    if (massArr[i] < massOcc && numMass[i] < numOfSlots) {
      Logger.log('Created option for ' + times[[i]])
      list.push(times[i] + ': ' + numMass[i] + ' of 23 slots filled');
    }
  }

  var item = form.getItems()[4];

  if (list.length != 0) {
    item.asListItem().setChoiceValues(list);
  } else {
    item.asListItem().setChoiceValues(['All Masses Full'])
  }
}

function emailReply() {
  var SHEET_ID = FormApp.getActiveForm().getDestinationId();
  var ss = SpreadsheetApp.openById(SHEET_ID);  //Responses spreadsheet with all masses
  var sheet = ss.getSheetByName('Master');  // Active Sheet

  var lastResp = sheet.getLastRow();
  var firstName = sheet.getRange(lastResp, 2).getValue();
  var email = sheet.getRange(lastResp, 4).getValue();
  var massAtt = sheet.getRange(lastResp, 6).getValue();

  var times = ['Saturday 5pm','Sunday 10am','Sunday 5pm','Sunday 9pm'];
  var mass = ''
  var isSat = false;
  for (var i = 0; i < times.length; i++) {
    if (massAtt.toString().includes(times[i])) {
      if (i == 0) {
        isSat = true;
      }
      mass = times[i];
      break;
    }
  }
  var sunday = new Date();
  var sat = new Date();
  sat.setDate(sat.getDate() + (6 - sat.getDay()));
  sunday.setDate(sunday.getDate() + (7 - sunday.getDay()));
  var subject = 'Mass Confirmation for ' + mass;

  var body = 'Hi ' + firstName + '!\n\nThank you for signing up for mass at the UW Catholic Newman Center!\nYou have been processed and are all set for the ' + mass + ' mass.\n\nSee you there!'

  Logger.log(email);
  Logger.log(subject);
  Logger.log(body);

  MailApp.sendEmail(email.toString().trim(), subject, body);
}



