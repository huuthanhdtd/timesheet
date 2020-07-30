// メッセージ送信
var Settings = {
  get: function (key) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("settings");
    var lastRow = sheet.getLastRow();
    var value = sheet.getRange(1, 1, lastRow, 2).getValues().filter(x => x[0] == key)

    if (value && value.length > 0)
      return value[0][1];

    return undefined;
  },
  
  set: function (key, value) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("settings");
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(1, 1, lastRow, 2).getValues() //.filter(x => x[0]==key)
    values.forEach(
      (row, i) => {
        if (row[0] == key) {
          sheet.getRange(i + 1, 2, 1, 1).setValue(value);
          return;
        }
      }
    );
    //return message;
  },
  
  getUser: function (username, datetime) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("worktimes");
    var labels = sheet.getRange("A1:Z2").getValues();
    var users = labels[0].map((e, i) => {
      if (e === "" || i == 0) return undefined;
      return { 'name': e, 'skype': labels[1][i] }
    });

    return users.filter(f => f !== undefined && f.name === username)[0];

  },
  
  getUsers: function (action, datetime) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("worktimes");
    var labels = sheet.getRange("A1:Z2").getValues();
    var users = labels[0].map((e, i) => {
      if (e === "" || i == 0) return undefined;
      return { 'name': e, 'skype': labels[1][i] }
    }).filter(e => e !== undefined);

    switch (action) {
      case 'signined':
        var firstDate = new Date(sheet.getRange("$A$3").getValue());
        var r = Math.floor((datetime - firstDate) / (24 * 3600 * 1000)) + 3;
        var signed = sheet.getRange("A" + r + ":Z" + r).getValues()[0].map((e, i) => ((i - 1) % 3 === 0 && e !== "") ? users[Math.floor((i - 1) / 3)] : undefined).filter(e => e != undefined);
        return signed;
        break;
      case 'notsignin':
        var firstDate = new Date(sheet.getRange("$A$3").getValue());
        var r = Math.floor((datetime - firstDate) / (24 * 3600 * 1000)) + 3;
        var notsigned = sheet.getRange("A" + r + ":Z" + r).getValues()[0].map((e, i) => ((i - 1) % 3 === 0 && e === "") ? users[Math.floor((i - 1) / 3)] : undefined).filter(e => e != undefined);
        return notsigned;
        break;
      case 'notsignout':
        var firstDate = new Date(sheet.getRange("$A$3").getValue());
        var r = Math.floor((datetime - firstDate) / (24 * 3600 * 1000)) + 3;
        var notsigned = sheet.getRange("A" + r + ":Z" + r).getValues()[0].map((e, i) => ((i - 1) % 3 === 1 && e === "") ? users[Math.floor((i - 1) / 3)] : undefined).filter(e => e != undefined);
        return notsigned;
        break;
      default:
        return users
        break;
    }

  }
}

//function createTrigger()
//{
//   // 毎日8時頃に出勤してるかチェックする
//    ScriptApp.newTrigger('confirmSignIn')
//      .timeBased()
//      .everyDays(1)
//      .atHour(8)
//      .create();
//
//    // 毎日22時頃に退勤してるかチェックする
//    ScriptApp.newTrigger('confirmSignOut')
//      .timeBased()
//      .everyDays(1)
//      .atHour(19)
//      .create();
//}

var Utils = {
  checkHoliday: function(today) {
    var holidays = Settings.get("休日");
    var wday = parseInt(Utilities.formatDate(today, "GMT+0700", "u"));
    var wdayStr = '月火水木金土日'.charAt(wday - 1);
    // 休日ならチェックしない
    if (holidays.indexOf(wdayStr) >= 0) return true;
    // 祝日を確認
    var dayStr = Utilities.formatDate(today, "GMT+0700", "yyyy/M/d");
    holidays = Settings.get("holidays");
    if (holidays.indexOf(dayStr) >= 0) return true;
    return false;
  },
  parseDate: function(str) {
    str = String(str || "").toLowerCase().replace(/[Ａ-Ｚａ-ｚ０-９]/g, function (s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
    var today = new Date();

    if (str.match(/(明日|tomorrow|tomorow)/i)) {
      return new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
    }

    if (str.match(/(今日|today)/i)) {
      return today;
    }

    if (str.match(/(昨日|yesterday)/i)) {
      return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
    }

    var reg = /((\d{4})[-\/年]{1}|)(\d{1,2})[-\/月]{1}(\d{1,2})/i;
    var matches = str.match(reg);
    if (matches) {
      var year = parseInt(matches[2], 10);
      var month = parseInt(matches[3], 10);
      var day = parseInt(matches[4], 10);
      if ((year == undefined) || year < 1970) {
        //
        if ((today.getMonth() + 1) >= 11 && month <= 2) {
          year = today.getFullYear() + 1;
        }
        else if ((today.getMonth() + 1) <= 2 && month >= 11) {
          year = today.getFullYear() - 1;
        }
        else {
          year = today.getFullYear();
        }
      }

      return new Date(year, month - 1, day);
    }

    return new Date();
  },
}
