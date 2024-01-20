function weather_description(valueMap) {

  var date = new Date(valueMap["日期"]);
  var time = new Date(valueMap["時間"]);
  var date_string = Utilities.formatDate(date, "GMT+8", "yyyy-MM-dd");
  // Logger.log(date_string);
  var datetime = new Date(date_string + "T" + Utilities.formatDate(time, "GMT+8", "HH:mm:ss"));
  // Logger.log(datetime);

  header = {
      'Authorization': CWA_API_TOKEN,
      'Accept': '*/*',
      'Connection': 'keep-alive',
      'User-Agent': "AppScript/5.0",
      'Cache-Control': 'no-cache',
      'Accept-Encoding': 'gzip',
      'Accept-Charset': 'ISO-8859-1,UTF-8;q=0.7,*;q=0.7',
      'Accept-Language': 'de,en;q=0.7,en-us;q=0.3'
  };

  // giving a array of time slot of 3 hours, and match time in time slot
  var slot = [ new Date(date_string + "T00:00"),
               new Date(date_string + "T03:00"),
               new Date(date_string + "T06:00"),
               new Date(date_string + "T09:00"),
               new Date(date_string + "T12:00"),
               new Date(date_string + "T15:00"),
               new Date(date_string + "T18:00"),
               new Date(date_string + "T21:00")];
  var timeFrom = '';
  var timeTo = '';
  for (var i = 1; i < slot.length; i++) {
    if (datetime < slot[i]) {
      format_str = "yyyy-MM-dd'T'HH:mm:ss";
      timeFrom = Utilities.formatDate(slot[i-1], "GMT+8", format_str);
      timeTo = Utilities.formatDate(slot[i], "GMT+8", format_str);
      // Logger.log(timeFrom);
      // Logger.log(timeTo);
      break;
    }
  }

  var url = "https://opendata.cwa.gov.tw/api/v1/rest/datastore/F-D0047-061?format=JSON&locationName=$locationName&timeFrom=$timeFrom&timeTo=$timeTo&elementName=WeatherDescription";
  url = url.replace('$locationName', "%E5%A4%A7%E5%AE%89%E5%8D%80"); // 大安區
  url = url.replace('$timeFrom', timeFrom);
  url = url.replace('$timeTo', timeTo);
  Logger.log(url);

  var opt = {
    'headers': header,
    'muteHttpExceptions': true
  };

  var response = UrlFetchApp.fetch(url, opt);
  if (response != undefined) {
    // parse CSV
    // Logger.log(response.getContentText());

    // Parse Json
    content = JSON.parse(response.getContentText());
    data = content['records']['locations'][0]['location'][0]['weatherElement'][0]['time'][0]['elementValue'][0]['value'];
    Logger.log(data);
  }

  return data;
}
