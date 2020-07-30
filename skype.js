// メッセージ送信
var Skype = {
  getToken: function () {
    var url = 'https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token';
    var clientId = "YOUR_CLIENT_ID";
    var clientSecret = "YOUR_CLIENT_SECRET";

    var options = {
      'method': 'post',
      'contentType': 'application/x-www-form-urlencoded',
      'payload': `grant_type=client_credentials&client_id=${clientId}&client_secret=${clientSecret}&scope=https://api.botframework.com/.default`
    };

    // Use the clientSecret to get the access token.
    var response = UrlFetchApp.fetch(url, options);
    var tokens = JSON.parse(response.getContentText());

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

    var send_options = {
      'method': "post",
      'contentType': 'application/json',
      'payload': JSON.stringify(data),
      'headers': {
        "Authorization": `Bearer ${token}`
      }
    };
    var response = UrlFetchApp.fetch(url, send_options);

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
        }
      });
    }

    return platform;
  }
}