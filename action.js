var Action = {
  // 出勤
  checkIn: function (data) {
    var platform = Skype.platform(data);
    if (platform === false || platform === 'Windows') {
      saveToSheet(data, new Date(data.timestamp), 'signin');
    } else {
      Skype.send('[ERR] @%s ' + platform + 'デバイスでチェックイン不可能!><', data);
    }
  },
  // 退勤
  checkOut: function (data) {
    saveToSheet(data, new Date(data.timestamp), 'signout');
  },
  // 休暇申請
  off: function (data) {
    var date = Utils.parseDate(data.text);
    saveToSheet(data, date, 'off');
  },
  // 休暇取消
  cancelOff: function (data) {
    var date = Utils.parseDate(data.text);
    saveToSheet(data, date, 'canceloff');
  },
  // 出勤中
  whoIsIn: function (data) {
    
  },
  // 休暇中
  whoIsOff : function (data) {
    
  },
  // 出勤していない人にメッセージを送る
  confirmSignIn: function (data) {
    var today = new Date(data.timestamp); //new Date(2020, 7-1, 29);
    if (Utils.checkHoliday(today)) return;

    var u = Settings.getUsers('notsignin', today);
    if (u && u.length > 0){
      var text = u.map(x => '<at id ="' + x.skype + '">@' + x.name + '</at>').join(", ");
      Skype.send('今日は休暇ですか？ ' + text, data);
    }
    
  },
  // 出勤していない人にメッセージを送る
  confirmSignOut: function (data) {
    var today = new Date(data.timestamp); //new Date(2020, 7-1, 29);
    if (Utils.checkHoliday(today)) return;

    var u = Settings.getUsers('notsignout', today);
    if (u && u.length > 0){
      var text = u.map(x => '<at id ="' + x.skype + '">@' + x.name + '</at>').join(", ");
      Skype.send('退勤しましたか？ ' + text, data);
    }
  },
  // ヘルプ
  help : function (data) {
    var help = '**Cú pháp cơ bản**\n```\n1. Checkin:  @OutOfBox おはよう|hi|hello|morning|出勤\n2. Checkout: @OutOfBox bye|night|お疲れ\n3. Đăng ký nghỉ\n  - Hôm nay: @OutOfBox off / @OutOfBox off today\n  - Ngày mai: @OutOfBox off tomorrow\n  - Ngày khác: @OutOfBox off yyyy/M/D\n4. Hủy đăng ký nghỉ\n  - Hôm nay: @OutOfBox cancel off / @OutOfBox cancel off today\n  - Ngày mai: @OutOfBox cancel off tomorrow\n  - Ngày khác: @OutOfBox cancel off yyyy/M/D\n```';
    Skype.send(help, data);
  },
}