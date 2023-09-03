//
// This is a Google app script for Google Spreadsheet
//

var LineNotifyEndPoint = "https://notify-api.line.me/api/notify";

//
// Send message to Line Notify
//

function sendLineNotification(token, message) {
    var formData = {
        "message": message
    };

    var options = {
        "headers": { "Authorization": "Bearer " + token },
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
