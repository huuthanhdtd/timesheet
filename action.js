var Action = {
  checkIn: function (data) {
    var platform = Skype.platform(data);
    if (platform === false || platform === 'Windows') {
      saveToSheet(data, new Date(data.timestamp), 'signin');
    } else {
      Skype.send('[ERR] @%s _Không thể check-in bằng thiết bị ' + platform + '!><_', data);
    }
  },

  checkOut: function (data) {
    saveToSheet(data, new Date(data.timestamp), 'signout');
  },

  off: function (data) {
    var date = Utils.parseDate(data.text);
    saveToSheet(data, date, 'off');
  },

  cancelOff: function (data) {
    var date = Utils.parseDate(data.text);
    saveToSheet(data, date, 'canceloff');
  },

  // 出勤していない人にメッセージを送る
  confirmSignIn: function (data) {
    var holidays = Settings.get("休日");
    var today = new Date(); //new Date(2020, 7-1, 29);
    var wday = parseInt(Utilities.formatDate(today, "GMT+0700", "u"));
    var wdayStr = '月火水木金土日'.charAt(wday - 1);
    // 休日ならチェックしない
    if (holidays.indexOf(wdayStr) >= 0) return;

    var u = Settings.getUsers('notsignin', today);
    if (u && u.length > 0){
      var text = u.map(x => '<at id ="' + x.skype + '">@' + x.name + '</at>').join(", ");
      Skype.send('今日は休暇ですか？ ' + text, data);
    }
    
  },

  // 出勤していない人にメッセージを送る
  confirmSignOut: function (data) {
    var holidays = Settings.get("休日");
    var today = new Date(); //new Date(2020, 7-1, 29);
    var wday = parseInt(Utilities.formatDate(today, "GMT+0700", "u"));
    var wdayStr = '月火水木金土日'.charAt(wday - 1);
    // 休日ならチェックしない
    if (holidays.indexOf(wdayStr) >= 0) return;

    var u = Settings.getUsers('notsignout', today);
    if (u && u.length > 0){
      var text = u.map(x => '<at id ="' + x.skype + '">@' + x.name + '</at>').join(", ");
      Skype.send('退勤しましたか？ ' + text, data);
    }
  },
}