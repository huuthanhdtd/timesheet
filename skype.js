// メッセージ送信
var Skype = {
  getToken: function () {
    // return 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Imh1Tjk1SXZQZmVocTM0R3pCRFoxR1hHaXJuTSIsImtpZCI6Imh1Tjk1SXZQZmVocTM0R3pCRFoxR1hHaXJuTSJ9.eyJhdWQiOiJodHRwczovL2FwaS5ib3RmcmFtZXdvcmsuY29tIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvZDZkNDk0MjAtZjM5Yi00ZGY3LWExZGMtZDU5YTkzNTg3MWRiLyIsImlhdCI6MTU5NTkyOTk0NCwibmJmIjoxNTk1OTI5OTQ0LCJleHAiOjE1OTYwMTY2NDQsImFpbyI6IkUyQmdZQWlOWlR2VSt0dFQvWW5TcEMvc1M2OWJBd0E9IiwiYXBwaWQiOiJkYzY5YjNmNC02M2QwLTQwMDctYmJlZC00Y2NjYjA3ZjE2ZGYiLCJhcHBpZGFjciI6IjEiLCJpZHAiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9kNmQ0OTQyMC1mMzliLTRkZjctYTFkYy1kNTlhOTM1ODcxZGIvIiwidGlkIjoiZDZkNDk0MjAtZjM5Yi00ZGY3LWExZGMtZDU5YTkzNTg3MWRiIiwidXRpIjoiOHZldDljMlp2a0NabkpTSlQwNW1BQSIsInZlciI6IjEuMCJ9.uwLLZufi2eUqBXSyvvJ0Umz7wz5ygKAMENS4d9iEZl_v8llDr7nOFlsKcenbUyN7xol0g3IwJQjQdlLEOsaBOhDkEgp95zJSZbejNX_fBUcUdg5wLmOojEeSF-FNhMU2mPLe8K53H2dt6uQI6t770n1w87OCB07uJYgGKsQitm0_t3P-9pw-6Z9KrkbD9N41Hztb0TyqNGI4EIwZCyN4w43Isk7CESzz7S9_q9kza4UR__-RxVf5Bj3-vLjG_AD1V8OG_OIwc0mJ9j6_jetn9hoyfjWdFD9QLtFs2vbjvToXfpBUzhs70-xOkZ7SCend7ccL61mjR1CTEUvAQKpKqQ';
    //    var token = Settings.get('token');
    //    if (token != '') return token;
    var url = 'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token';
    var clientId = "dc69b3f4-63d0-4007-bbed-4cccb07f16df";
    var clientSecret = "7pNlxU73Og24Kq.pso8VJT_1RSf_Fs~5GG";

    var options = {
      'method': 'post',
      'contentType': 'application/x-www-form-urlencoded',
      'payload': `grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}&scope=https://api.botframework.com/.default`
    };


    // Use the clientSecret to get the access token.
    var response = UrlFetchApp.fetch(url, options);
    var tokens = JSON.parse(response.getContentText());
    /*
    // save to sheet
    var book = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = book.getSheetByName("settings");
    sheet.getRange("B1").setValue(tokens.access_token);
    */
    //Settings.set('token', tokens.access_token);
    return tokens.access_token;

  },
  send: function (message, options) {
    var url = `${options.serviceUrl}/v3/conversations/${options.conversation.id}/activities/${options.id}`;
    var token = Settings.get('token'); // this.getToken();
    if (token == '') return false;
    var data = {
      "type": "message",
      //    "from": {
      //      "id": options.from.id,
      //      "name": "botName"
      //    },
      //    "conversation": {
      //      "id": options.conversation.id,
      //      "name": "conversationName"
      //    },
      //    "recipient": {
      //      "id": options.recipient.id,
      //      "name": options.recipient.name
      //    },
      "text": Utilities.formatString(message, options.from.name), //"Chào " + options.from.name + ". " + message, // "Chúc bạn làm việc hiệu quả!^^;",
      //"replyToId": options.recipient.id
    }
    var payload = JSON.stringify(data);

    var send_options = {
      'method': "post",
      'contentType': 'application/json',
      'payload': payload,
      'headers': {
        "Authorization": `Bearer ${token}`
      }
    };
    //MailApp.sendEmail("nht@huuthanhdtd.com", "send_options", JSON.stringify(send_options));

    var response = UrlFetchApp.fetch(url, send_options);

    //MailApp.sendEmail("nht@huuthanhdtd.com", "response", response.getContentText());



    return message;
  },
  platform: function (data) {
    // Check platform 
    var platform = false;
    if (data.entities) {
      data.entities.forEach(entry => {
        if (entry.platform) {
          platform = entry.platform;
          return;
          //if (entry.platform === "Windows") { return;}
        }
      });

    }

    return platform;
  }
}