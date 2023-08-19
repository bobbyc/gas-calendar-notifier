//
// This is a Google app script for Google Spreadsheet
//

var LineNotifyEndPoint = "https://notify-api.line.me/api/notify";
var AccessToken = ""; // Testing
// var AccessToken = ""; // 北小泳隊

//
// Send message to Line Notify
//

function test_notification() {
  sendLineNotification("麥克風測試 1, 2, 3");
}

function sendLineNotification(message) {
    var formData = {
        "message": message
    };

    var options = {
        "headers": { "Authorization": "Bearer " + AccessToken },
        "method": 'post',
        "payload": formData
    };

    try {
        var response = UrlFetchApp.fetch(LineNotifyEndPoint, options);
    }
    catch (error) {
        Logger.log(error.name + ":" + error.message);
        return;
    }

    if (response.getResponseCode() !== 200) {
        Logger.log("Error: " + response.getContentText());
    }
}
