// Skypeから受信データ
function receiveMessage(data) {
  // コマンド集
  var commands = [
    ['checkOut', /(bye|night|バ[ー〜ァ]*イ|ば[ー〜ぁ]*い|おやすみ|お[つっ]ー|おつ|さらば|お先|お疲|帰|乙|bye|night|(c|see)\s*(u|you)|退勤|ごきげんよ|グ[ッ]?バイ|tạm biệt)/i],
    //      ['actionWhoIsOff', /(だれ|誰|who\s*is).*(休|やす(ま|み|む))/i],
    //      ['actionWhoIsIn', /(だれ|誰|who\s*is)/i],
    ['cancelOff', /((休|やす(ま|み|む)|休暇).*(キャンセル|消|止|やめ|ません)|cancel off|cancel)/i],
    ['off', /(休|やす(ま|み|む)|休暇|off)/i],
    ['checkIn', /(chào|おはよう|モ[ー〜]+ニン|も[ー〜]+にん|おっは|おは|へろ|はろ|ヘロ|ハロ|hi|hello|morning|出勤)/i],
    ['confirmSignIn', /__confirmSignIn__/i],
    ['confirmSignOut', /__confirmSignOut__/i],
  ];

  // メッセージを元にメソッドを探す
  var command = commands.find(e => data.text.match(e[1]));

  // メッセージを実行
  if (command && command[0]) {
    return Action[command[0]](data);
  }
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  if (data.type === "message") {
    receiveMessage(data)
  }

  //  MailApp.sendEmail("nht@huuthanhdtd.com", "skype bot test",
  //                      JSON.stringify(data));
}


function saveToSheet(data, datetime, action) {
  var name = data.from.name;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("worktimes");
  //var array = sheet.getRange("A1:A").getValues();

  var dateStr = Utilities.formatDate(datetime, "GMT+0700", "YYYY/M/d");
  var firstDate = new Date(sheet.getRange("$A$3").getValue());
  //find row
  var r = Math.floor((datetime - firstDate) / (24 * 3600 * 1000)) + 3;
  //save current date
  if (!sheet.getRange(r, 1, 1, 1).getValue()) sheet.getRange(r, 1, 1, 1).setValue(dateStr);

  //find column
  var header = sheet.getRange("A1:Z1").getValues()[0];

  var c = header.indexOf(name) + 1;
  if (c <= 0) {
    var cnt = header.filter(y => y !== "").length - 1;
    c = cnt * 3 + 1 + 1;
    sheet.getRange(1, c, 1, 1).setValue(name);
  } else {

  }
  if (c > 0) {
    var skypeId = sheet.getRange(2, c, 1, 1).getValue();
    switch (action) {
      case 'signin':
        if (!sheet.getRange(r, c, 1, 1).getValue()) {
          var sChar = String.fromCharCode(65 + c - 1);
          var eChar = String.fromCharCode(65 + c);
          var fomular = '=if(or(' + sChar + r + '="";' + eChar + r + '="");"";' + eChar + r + '-' + sChar + r + ')'
          var row = [datetime, , fomular];
          sheet.getRange(r, c, 1, row.length).setValues([row]);

          Skype.send('<at id="' + skypeId + '">@%s</at> おはようございます!^^;', data);
        } else {
          Skype.send('<at id="' + skypeId + '">@%s</at>. 既にチェックインしました!^^;', data);
        }
        break;
      case 'signout':
        sheet.getRange(r, c + 1, 1, 1).setValue(datetime);
        Skype.send('<at id="' + skypeId + '">@%s</at>さん、お疲れ様でした!^^;', data);
        break;
      case 'off':
        if (!sheet.getRange(r, c, 1, 1).getValue()) {
          var row = ['-', '-', 'off'];
          sheet.getRange(r, c, 1, row.length).setValues([row]);
          Skype.send('<at id="' + skypeId + '">@%s</at> ' + dateStr + 'を休暇として登録しました', data);
        } else {
          Skype.send('[ERR] <at id="' + skypeId + '">@%s</at> チェックインしましたので*' + dateStr + '*の休暇を登録できません ', data);
        }
        break;
      case 'canceloff':
        var row = sheet.getRange(r, c, 1, 3).getValues()[0]
        if (row[0] == '-' && row[2] == 'off') {
          var row = [null, null, action];
          sheet.getRange(r, c, 1, row.length).setValues([row]);
          Skype.send('<at id="' + skypeId + '">@%s</at> *' + dateStr + '*の休暇を取り消しました', data);
        } else {
          Skype.send('[INF] <at id="' + skypeId + '">@%s</at> *' + dateStr + '*の休暇を登録していません', data);
        }
        break;
      default:
        break;
    }

  }

}


// Time-based triggerで実行
function confirmSignIn() {
  var x = {
    "text": "__confirmSignIn__",
    "type": "message",
    "timestamp": "2020-07-28T10:26:53.34Z",
    "id": "***",
    "channelId": "skype",
    "serviceUrl": "https://smba.trafficmanager.net/apis/",
    "from": {
      "id": "29:1gIByJgqnMHl69AY6CatsTkxOLWshj71L_L_cjNQVHpw",
      "name": "Nguyen Huu THANH"
    },
    "conversation": {
      "isGroup": true,
      "id": "19:6dcea85b777b4c3faaaefeab9901d004@thread.skype"
    }
  }
  receiveMessage(x);
}

// Time-based triggerで実行
function confirmSignOut() {
  var x = {
    "text": "__confirmSignOut__",
    "type": "message",
    "timestamp": "2020-07-28T10:26:53.34Z",
    "id": "***",
    "channelId": "skype",
    "serviceUrl": "https://smba.trafficmanager.net/apis/",
    "from": {
      "id": "29:1gIByJgqnMHl69AY6CatsTkxOLWshj71L_L_cjNQVHpw",
      "name": "Nguyen Huu THANH"
    },
    "conversation": {
      "isGroup": true,
      "id": "19:6dcea85b777b4c3faaaefeab9901d004@thread.skype"
    }
  }
  receiveMessage(x);
  // 毎日tokenを更新するため
  Settings.set('token', Skype.getToken());
}


function testAction() {
  var x = {
    "text": "__confirmSignOut__",
    "type": "message",
    "timestamp": "2020-07-30T09:24:55.972Z",
    "id": "***",
    "channelId": "skype",
    "serviceUrl": "https://smba.trafficmanager.net/apis/",
    "from": {
      "id": "***",
      "name": "フィタン(Stellar Doh)"
    },
    "entities": [
      //      {
      //        "mentioned":{
      //          "id":"28:dc69b3f4-63d0-4007-bbed-4cccb07f16df"
      //        },
      //        "text":"<at id=\"28:dc69b3f4-63d0-4007-bbed-4cccb07f16df\">OutOfBox</at>",
      //        "type":"mention"
      //      },
      //      {
      //        "locale":"ja-JP",
      //        "country":"VN",
      //        "platform":"Windows",
      //        "timezone":"Asia/Bangkok",
      //        "type":"clientInfo"
      //      }
    ],
    "conversation": {
      "id": "29:1lSB24MkgCKlnL7qG-9rpwAgGzcCvx3l6QJFKj6ypzXQ"
    }
  }

  receiveMessage(x);
}
