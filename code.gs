// Slackのチャンネル一覧とpurposeを投稿
//[ファイル]→[プロジェクトのプロパティ]→[スクリプトのプロパティ]に以下の各プロパティを設定すること
// slack_api_token  APIトークン。 https://api.slack.com/docs/oauth-test-tokens で Create Token を押すと生成される xxxx-1234567890-1234567890 のような値
// channel          投稿するチャンネル

function slackApi(token, path, params) {
  if(!params) params = {};
  var url = "https://slack.com/api/" + path;
  var q = "?token=" + encodeURIComponent( token );
  for( var key in params ){
    q += "&" + encodeURIComponent( key ) + "=" + encodeURIComponent( params[ key ] );
  }
  url += q;
  var res = UrlFetchApp.fetch( url );
  var ret = JSON.parse( res.getContentText() );
  if( ret.error ){
    throw "GET " + path + ": " + ret.error;
  }
  return ret;
}

function slackPost (token, channel, text, attachments) {
  var options = {
    "method" : "POST",
    "payload" : {
      "channel" : channel,
      "text" : text,
      "username" : "bot"
    }
  };
  if( attachments !== undefined ){
    if( !attachments instanceof Array ){
      attachments = [ attachments ];
    }
    options.payload.attachments = JSON.stringify( attachments );
  }
  return UrlFetchApp.fetch("https://slack.com/api/chat.postMessage?token=" + token, options);
}

function main(){
  var channels, token, channel, text = "";
  var values = [['チャンネル名', 'チャンネル概要', '参加人数', 'リーダー', '副リーダー']];
  var i;
  var date_obj = new Date();
  date = Utilities.formatDate(date_obj, 'Asia/Tokyo', 'yyyy-M-d');
  
  // チャンネル一覧の取得
  token = PropertiesService.getScriptProperties().getProperty("slack_api_token");
  channel = PropertiesService.getScriptProperties().getProperty("channel");
  if (channel.charAt(0) !== "#") channel = "#" + channel;
  if (!token) {
    throw 'You should set "slack_api_token" property from [File] > [Project properties] > [Script properties]';
  }
  channels = slackApi(token, "channels.list");
  channels.channels.forEach( function (ch) {
    if (!ch.is_archived) {
      // 副リーダーは'-'と表示したいので分岐させている
      if (ch.name.match(/project-/)) {
        text += "#" + ch.name + " " + ch.purpose.value + "\n";
        values.push(["#" + ch.name, ch.purpose.value, ch.num_members, '', '-']);
      }
      else if (ch.name.match(/team-/)) {
        text += "#" + ch.name + " " + ch.purpose.value + "\n";
        values.push(["#" + ch.name, ch.purpose.value, ch.num_members, '', '']);
      }
    }
  });
  
  // チャンネル一覧をスプレッドシートに挿入
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.insertSheet(date, 0)
  var sheet = spreadsheet.getActiveSheet();
  var range = sheet.getRange(1, 1, values.length, values[0].length);
  for (i = 1; i <= values[0].length; i++) {
    sheet.setColumnWidth(i, 250);
  }
  range.setWrap(true);
  range.setValues(values);
  
  var numColumns = sheet.getLastColumn();
  for (i = 1; i <= sheet.getLastRow(); i++){
    if (i == 1) {
      sheet.getRange(i, 1, 1, numColumns).setBackgroundColor('#BDBDBD');
    } else if (i % 2 == 0) {      
      sheet.getRange(i, 1, 1, numColumns).setBackgroundColor('#F3F3F3');    
    } else {
      // 白に塗りつぶされているので何もしない
    }
  }

  // slackにチェンネル一覧をポストする
  // slackPost(token, channel, text);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();           // Uiクラスを取得する
  var menu = ui.createMenu('スクリプト');  // Uiクラスからメニューを作成する
  menu.addItem('チャンネル一覧の取得', 'main');   // メニューにアイテムを追加する
  menu.addToUi();                            // メニューをUiクラスに追加する
}
