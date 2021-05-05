function organize() {
  var formSS = SpreadsheetApp.getActive();
  var formSheet = formSS.getSheetByName('Master');
  var ezSheet = formSS.getSheetByName('EZ View Sheet');

  var numResp = formSheet.getLastRow() - 1;
  var timestamp = formSheet.getRange(2, 1, numResp).getValues();
  var firstName = formSheet.getRange(2, 2, numResp).getValues();
  var lastName = formSheet.getRange(2, 3, numResp).getValues();
  var massCount = formSheet.getRange(2, 5, numResp).getValues();
  var massAtt = formSheet.getRange(2, 6, numResp).getValues();
  var today = new Date();
  var startOfWeek = new Date();
  var days = today.getDay();
  if (days == 0) {
    days = 7;
  }
  startOfWeek.setDate(today.getDate() - (days - 1));
  var months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var times = ['Saturday 5pm','Sunday 10am','Sunday 5pm','Sunday 9pm'];
  var list = [[], [], [], []]
  Logger.log(timestamp);

  for (var i = 0; i < firstName.length; i++) {
    var data = '' + firstName[i] + ' ' + lastName[i];
    var resp = massAtt[i];
    var currDate = timestamp[i].toString().split(' ');
    Logger.log(currDate);
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
        if (!ezSheet.getRange(3, 2, numResp).getValues().toString().includes(data)) {
          list[0].push(data + ' ' + massCount[i]);
        }
      } else if (resp.toString().search(times[1]) != -1) {
        if (!ezSheet.getRange(3, 4, numResp).getValues().toString().includes(data)) {
          list[1].push(data + ' ' + massCount[i]);
        }
      } else if (resp.toString().search(times[2]) != -1) {
        if (!ezSheet.getRange(3, 6, numResp).getValues().toString().includes(data)) {
          list[2].push(data + ' ' + massCount[i]);
        }
      } else if (resp.toString().search(times[3]) != -1) {
        if (!ezSheet.getRange(3, 8, numResp).getValues().toString().includes(data)) {
          list[3].push(data + ' ' + massCount[i]);
        }
      }
    }
  }
  Logger.log(list);


  var maxRow = Math.max(list[0].length, list[1].length, list[2].length, list[3].length);

    if (maxRow == 0) {
    return;
  }

  for (var j = 0; j < maxRow; j++) {
    var temp = [''];
    for (var k = 0; k < 4; k++) {
      if (j >= list[k].length) {
        temp.push('');
      } else {
        temp.push(list[k][j]);
      }
      temp.push('');
    }
    ezSheet.appendRow(temp);
  }
}

function autoSort() {
  var formSS = SpreadsheetApp.getActive();
  var ezSheet = formSS.getSheetByName('EZ View Sheet');
  var emailBank = formSS.getSheetByName('Email Bank');
  var numEmails = emailBank.getLastRow()
  emailBank.getRange(2, 1, numEmails, 3).sort(1);
  var numResp = ezSheet.getLastRow();
  if (numResp == 2) {
    return;
  }
  for (var i = 2; i < 10; i += 2) {
    var names = ezSheet.getRange(3, i, numResp - 2, 2).sort(i);
  }
}

function attendence() {
  var formSS = SpreadsheetApp.getActive();
  var ezSheet = formSS.getSheetByName('EZ View Sheet');
  var emailSheet = formSS.getSheetByName('Email Bank');
  for (var i = 2; i < 9; i += 2) {
    var rowNum = ezSheet.getLastRow() - 2;
    if (rowNum == 0) {
      return;
    }
    var dataCol = ezSheet.getRange(3, i, rowNum).getValues();
    var attenCol = ezSheet.getRange(3, i + 1, rowNum).getValues();
    var topCell = ezSheet.getRange(1, i+1);
    var total = 0;
    var sum = 0;
    for (var j = 0; j < dataCol.length; j++) {
      if (dataCol[j][0] === '') {
        break;
      }
      
      if (attenCol[j][0].toString().includes('@')) {
        var email = attenCol[j][0].toString();
        attenCol[j][0] = 'w';
        ezSheet.getRange(3, i+1, rowNum).setValues(attenCol);
        var name = dataCol[j][0];
        name = name.toString().split(' ');
        var emailDatabase = emailSheet.getRange(2, 3, emailSheet.getLastRow()).getValues().toString();
        if (name.length >= 2 && !emailDatabase.includes(email)) {
          emailSheet.appendRow([name[0], name[name.length - 2], email]);
        }
      }
      if (attenCol[j][0] !== 'w') {
        var num = dataCol[j][0].toString().trim().split(' ');
        total += parseInt(num[num.length - 1]);
      }
      if (attenCol[j][0] !== '') {
        var num = dataCol[j][0].toString().trim().split(' ');
        sum += parseInt(num[num.length - 1]);
      }
    }
    topCell.setValues([['' + sum + '/' + total]]);
  }
}

function emailArchiver() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();  //Responses spreadsheet with all masses
  var sheet = ss.getSheetByName('Master');  // Active Sheet

  var numResp = sheet.getLastRow();
  var emailData = sheet.getRange(2, 4, numResp).getValues();
  var firstName = sheet.getRange(2, 2, numResp).getValues();
  var lastName = sheet.getRange(2, 3, numResp).getValues();

  
  // Email bank stuff
  var emailSheet = ss.getSheetByName('Email Bank');
  var emailBankInd = emailSheet.getLastRow();
  var emailBank = emailSheet.getRange(2, 3, emailBankInd).getValues();
  //Logger.log(emailData);

  for (var i = 0; i < emailData.length; i++) {
    var email = emailData[i].toString();
    //Logger.log('Current Email: ' + email);
    var cont = false;
    for (var j = 0; j < emailBank.length; j++) {
      if (email === emailBank[j].toString()) {
        cont = true;
        break;
      }
    }
    if (!cont) {
      // Add email to sheet
      if (email.toString().includes('@')) {
        emailSheet.appendRow([firstName[i].toString(), lastName[i].toString(), email.toString()]);
        Logger.log('Email added : ' + email);
      }
      emailBankInd = emailSheet.getLastRow();
    }
  }
}

function ezWipe() {
  var formSS = SpreadsheetApp.getActive();
  var ezSheet = formSS.getSheetByName('EZ View Sheet');
  ezSheet.clearContents();
  ezSheet.appendRow([' ']);
  ezSheet.appendRow([' ', 'Sat 5pm',	'Attended',	'Sun 10am',	'Attended',	'Sun 5pm',	'Attended',	'Sun 9pm', 'Attended'])
}
