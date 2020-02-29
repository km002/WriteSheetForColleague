//LINE developersのメッセージ送受信設定に記載のアクセストークン
var ACCESS_TOKEN = '<アクセストークン>';
var ownerID = ""; // メールアドレス
var templateid = "<出勤簿のテンプレートとして使うスプレッドシートのid>"; // 出勤簿テンプレート
//var taikinTemplate = ""; // db用シート名　
var destfolderid = ""; // 保存用フォルダディレクトリid
var db = '開発者がdbとして使うスプレッドシート'; // 開発者のdb用スプレッドシート　
var userName = "";
var userNo = "";
var userId = "";

function doPost(e) {

	// WebHookで受信した応答用Token
    var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
//	var replyToken = ""; //GASのデバッグ機能使用時に使う
    Logger.log(replyToken);

    // ユーザIDを取得
//	var userId = "";
      userId = JSON.parse(e.postData.contents).events[0].source.userId;

// ユーザーのメッセージを取得
	var userMessage = JSON.parse(e.postData.contents).events[0].message.text;

	// 応答メッセージ用のAPI URL
	var url = 'https://api.line.me/v2/bot/message/reply';

	// DB用シートを取得
	var template = DriveApp.getFileById(db);

	var taikinSheet = SpreadsheetApp.open(template);
	var taikinSheets = taikinSheet.getSheets();
	Logger.log("Sheet名:" + taikinSheets[0].getSheetName());
	if ( taikinSheets[0].getSheetName() == "シート1" ) {
		var userInfo = taikinSheets[0];
	}
	// ユーザからの入力をDBに転記
	userInfo.getRange(1, 16).setValue(userMessage);
	var memo = userInfo.getRange(1,16).getValue();
    var dateMemo = memo;
    if(Object.prototype.toString.call(memo) == "[object Date]"){
      memo = memo.toString();
    }

	// LINEからの初期設定受付
	// メール
	if(memo.indexOf('@gmail.com') != -1){
		userInfo.getRange(4, 16).setValue('分岐1-1');
		ownerID = memo.substring(0, memo.indexOf('@')+10);
		userName = memo.substring(memo.indexOf('@')+10);
        if(userName.match(/ /)){
          userName = userName.substring(1);
        }
//        userName = "";
		var userRange = '';
		for(var i = 1;i < 30;i++){
			userRange = userInfo.getRange(i,1).getValue();
            userIdRange = userInfo.getRange(i,2).getValue();
            if(userIdRange == userId){
              	userNo = i;
				userInfo.getRange(i,1).setValue(ownerID);
				userInfo.getRange(i,2).setValue(userId);
				userInfo.getRange(i,4).setValue(userName);
                break;
            }
			if(userRange == ""){
				userNo = i;
				userInfo.getRange(i,1).setValue(ownerID);
				userInfo.getRange(i,2).setValue(userId);
				userInfo.getRange(i,4).setValue(userName);
                break;
			}
		}
		Logger.log(userInfo.getRange(i,1).getValue);
		Logger.log(userInfo.getRange(i,2).getValue);
		Logger.log(userInfo.getRange(i,4).getValue);
        userInfo.getRange(2,17).setValue("@gmail.com");
		// シートで確認するログ
		userInfo.getRange(5,16).setValue(userInfo.getRange(i,1).getValue());
//		userInfo.getRange(6,16).setValue(userInfo.getRange(i,2).getValue());
		userInfo.getRange(7,16).setValue(userInfo.getRange(i,4).getValue());
		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'Gmailと名前を設定しました。\nGmail:' + userInfo.getRange(userNo,1).getValue() + "\n名前:" + userInfo.getRange(userNo,4).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}else if(memo.match(/folders/) || memo.match(/folderview?/) || memo.match(/open?/)){
		// userNoを取得
		for(var i = 1;i < 30;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
        userInfo.getRange(2,17).setValue("フォルダ");
		//フォルダID
        if(memo.match(/folders/)){
        destfolderid = memo.substring(memo.indexOf("folders/")+8);
        } else {
        destfolderid = memo.substring(memo.indexOf("=")+1);
        }
		userInfo.getRange(userNo,3).setValue(destfolderid);

        if(memo.match(/sharing/) || memo.match(/folderview?/) || memo.match(/open?/)){
		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'フォルダIDを設定しました。\n出勤簿は下記のフォルダで作成します。\n' + "https://drive.google.com/drive/folders/" + userInfo.getRange(userNo,3).getValue(),
				}],
			}),
		});
        } else {
        		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'リンクが共有可能ではありません。\nもう1度「共有可能なリンクを取得」から取得したリンクを教えてください。\n',
				}],
			}),
		});
        }
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}
	// 定時
	else if(memo.match(/定時/)){
		// userNoを取得
		for(var i = 1;i < 30;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
        userInfo.getRange(2,17).setValue("定時");
		//出勤（時）
		userInfo.getRange(userNo,5).setValue(memo.substring(2,4));
		//出勤（分）
		userInfo.getRange(userNo,6).setValue(memo.substring(4,6));
		//退勤（時）
		userInfo.getRange(userNo,7).setValue(memo.substring(6,8));
		//退勤（分）
		userInfo.getRange(userNo,8).setValue(memo.substring(8,10));

		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'定時を設定しました。\n出勤:' + userInfo.getRange(userNo,5).getValue() + '時' + userInfo.getRange(userNo,6).getValue() + '分' + '\n退勤:' + userInfo.getRange(userNo,7).getValue() + '時' + userInfo.getRange(userNo,8).getValue() + '分',
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
	} else if(Object.prototype.toString.call(dateMemo) == "[object Date]"){
        userInfo.getRange(2,17).setValue("退勤到達");
		// userNoを取得
		for(var i = 1;i < 30;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		// ユーザーの入力した退勤時間を設定
		userInfo.getRange(userNo,9).setValue(dateMemo);
		// 退勤時間を記録する出勤簿を設定
		var file = SpreadsheetApp.openByUrl(userInfo.getRange(userNo,12).getValue());
		// メッセージを日付と時刻に分解
		var year = dateMemo.getFullYear();
		var month = dateMemo.getMonth();
		var date = dateMemo.getDate();
		var hour = dateMemo.getHours();
		var minutes = dateMemo.getMinutes();
		if(minutes < 30){
			minutes = 0;
		} else if(minutes > 30){
			minutes = 30;
		}
        userInfo.getRange(2,17).setValue("退勤到達2");

		var message = date + '日は' + hour + '時' + minutes + '分';
        
        userInfo.getRange(2,17).setValue("退勤到達3");

        userInfo.getRange(3,17).setValue(file);
        userInfo.getRange(4,17).setValue(date);
        userInfo.getRange(5,17).setValue(hour);
        userInfo.getRange(6,17).setValue(minutes);
		var spreadsheet = writeTaikin(file,date,hour,minutes);
        
        userInfo.getRange(2,17).setValue("退勤到達4");
		// ログ
		userInfo.getRange(3, 16).setValue(spreadsheet.getUrl());
        
        userInfo.getRange(2,17).setValue("退勤到達5");

		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':message + 'に退勤ですね。\n退勤時間を記録しました。\n' + spreadsheet.getUrl(),
				}],
			}),
		});


		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}
	else if (memo.match(/spreadsheets/) && memo.match(/edit/)){
		// userNoを取得
		for(var i = 1;i < 30;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		// ログ
		userInfo.getRange(4, 16).setValue('出勤簿設定分岐');
		// 出勤簿を設定
		userInfo.getRange(userNo,12).setValue(memo);

		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'退勤時間を記録する出勤簿を設定しました。\n' + userInfo.getRange(userNo,12).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}
	else if (memo.match(/出勤簿/) && memo.match(/作/) && memo.match(/先月/) == null){
		// userNoを取得
		for(var i = 1;i < 230;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		// ログ
		userInfo.getRange(4, 16).setValue('分岐3');
		//日付作成
		var now = new Date();
		var month = now.getMonth()+1;    //月
		// 出勤簿作成
		var newfile = writeSheet(month,userNo);
		// 作成した最新の出勤簿を記録
		userInfo.getRange(userNo,12).setValue(newfile.getUrl())

		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'出勤簿を作成しました。\n' + userInfo.getRange(userNo,12).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	} else if (memo.match(/先月/) && memo.match(/出勤簿/) && memo.match(/作/)){
		// userNoを取得
		for(var i = 1;i < 30;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		//日付作成
		var now = new Date();
		var month = now.getMonth();    //月
		// 出勤簿作成
		var newfile = writeSheet(month,userNo);
		// 作成した最新の出勤簿を記録
		userInfo.getRange(userNo,12).setValue(newfile.getUrl())

		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'出勤簿を作成しました。\n' + userInfo.getRange(userNo,12).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}else{

        		UrlFetchApp.fetch(url, {
			'headers': {
				'Content-Type': 'application/json; charset=UTF-8',
				'Authorization': 'Bearer ' + ACCESS_TOKEN,
			},
			'method': 'post',
			'payload': JSON.stringify({
				'replyToken': replyToken,
				'messages': [{
					'type': 'text',
					'text':'できること一覧です。（●は初期設定項目です）\n●メール(Gmail)、名前設定(出勤簿記入用)\n(入力例「test@gmail.com テスト太郎」)\n●フォルダ設定\nGoogleDriveで共有可能にしたフォルダのurlを入力してください。参考 https://support.google.com/drive/answer/7166529\n●定時設定\n入力例「定時08301730」\n・出勤簿作成（処理に1分前後かかります）\n入力例「出勤簿作成」「出勤簿作って」など\n・先月出勤簿作成（処理に1分前後かかります）\n入力例「先月出勤簿作成」「先月の出勤簿作って」など\n・出勤簿設定\n（直近に作成した出勤簿以外の出勤簿に退勤時間を記録したい場合、使いたいスプレッドシートのurlを貼り付けてください）\n・退勤時間記録\n入力例「2020/02/28 11:09」\n(Google日本語入力等で「今日 今」と入力し変換してください)\n※使用上の注意※\n出勤簿をxlsx形式で保存しExcelで開くと非表示になっているシート「設定用」が表示されてしまいます。右クリック等で非表示にしてください。',
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}


}


function writeSheet(m,userNo) {
// DB用シートを取得
    var dbtemplate = DriveApp.getFileById(db); //下とかぶるのでdbtemplateにする

	var taikinSheet = SpreadsheetApp.open(dbtemplate);
	var taikinSheets = taikinSheet.getSheets();
	Logger.log("Sheet名:" + taikinSheets[0].getSheetName());
	if ( taikinSheets[0].getSheetName() == "シート1" ) {
		var userInfo = taikinSheets[0];
	}
    userInfo.getRange(2,18).setValue("writeSheet到達");
    userInfo.getRange(3,18).setValue("writeSheet到達2");
	// 保存フォルダを設定
    var sharing = userInfo.getRange(userNo,3).getValue().indexOf("?usp=sharing");
    if(sharing != -1){
      userInfo.getRange(6,18).setValue(sharing);
      var folderid = userInfo.getRange(userNo,3).getValue();
      destfolderid = folderid.substring(0, sharing);
    } else {
      var folderid = userInfo.getRange(userNo,3).getValue();
      destfolderid = folderid;
    }
    userInfo.getRange(6,18).setValue(destfolderid);
	userName = userInfo.getRange(userNo,4).getValue();
    userInfo.getRange(4,18).setValue("writeSheet到達3");
	//出勤簿テンプレートを指定
	var templateid = "163CopCFxrMshO8bYfqoY4Jlc3Rr9K-20_tn9njpAYwU";
	var template = DriveApp.getFileById(templateid);
    userInfo.getRange(5,18).setValue("writeSheet到達4");
	//保存フォルダを指定
	var destfolder = DriveApp.getFolderById(destfolderid);
    userInfo.getRange(6,18).setValue("writeSheet到達5");
	//日付作成
	var now = new Date();
	var year = now.getFullYear(); //年
	var month = m;    //月

	//テンプレートを元に出勤簿を作成
	if ( month.length === 2){
		var filename = "出勤簿（" + year + month + "）"+ userName +".xlsx";
	} else if (month === 0){
		var filename = "出勤簿（" + year + "01" + "）"+ userName +".xlsx";
	}
	else {
		var filename = "出勤簿（" + year + 0 + month + "）"+ userName +".xlsx";
	}
	Logger.log(filename);
	var newfile = template.makeCopy(filename, destfolder);

	//ファイルの共有設定（※権限付与はまた別途）
	newfile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

	// スプレッドシートのデータを得る
	var spreadsheet = SpreadsheetApp.open(newfile);
	var sheets = spreadsheet.getSheets();
	Logger.log("Sheet名:" + sheets[1].getSheetName());
	if ( sheets[1].getSheetName() === "原本" ) {
		var sheet = sheets[1];
		//シートの最終行番号、最終列番号を取得
		var startrow = 1;
		var startcol = 1;
		var lastrow = sheet.getLastRow();
		var lastcol = sheet.getLastColumn();


		//月を記入
		if (month === 0){
			sheet.getRange("D1").setValue(1);
		} else {
			sheet.getRange("D1").setValue(month);
		}
		Logger.log("月:" + sheet.getRange("D1").getValue());
		// 日数を取得
		var lastDay = getLastDay(year, month);
		Logger.log("日数" + lastDay);

		for (var j =5; j < lastDay+5;j++){
			//土日祝日は記入しない
			if(!(sheet.getRange(j, 2).getValue() === 7) && !(sheet.getRange(j, 2).getValue() === 1) && !(sheet.getRange(j, 2).getValue() === "祝")){
				var dateCheck = sheet.getRange(j, 1).getValue();
				Logger.log("日付:" + dateCheck.getDate());
				Logger.log("曜日:" + sheet.getRange(j, 2).getValue());
				sheet.getRange(j, 4).setValue(userInfo.getRange(userNo,5).getValue());
				sheet.getRange(j, 5).setValue(userInfo.getRange(userNo,6).getValue());
				sheet.getRange(j, 6).setValue(userInfo.getRange(userNo,7).getValue());
				sheet.getRange(j, 7).setValue(userInfo.getRange(userNo,8).getValue());
                //定時ログ
                userInfo.getRange(3,18).setValue("writeSheet到達");
                userInfo.getRange(4,18).setValue("writeSheet到達");
                userInfo.getRange(5,18).setValue("writeSheet到達");
                userInfo.getRange(6,18).setValue("writeSheet到達");
				sheet.getRange(j, 8).setValue(01);
				sheet.getRange(j, 9).setValue(00);
				sheet.getRange(j, 10).setValue(00);
				sheet.getRange(j, 11).setValue(00);
			} else {
				Logger.log("書き込まない曜日:" + sheet.getRange(j, 2).getValue());
			}
		}

		//名前を更新
		sheet.getRange("AE1").setValue(userName);



	}
	return spreadsheet;
}


function writeTaikin(file, date, hour, minutes){
// DB用シートを取得
    var dbtemplate = DriveApp.getFileById(db); //下とかぶるのでdbtemplateにする

	var taikinSheet = SpreadsheetApp.open(dbtemplate);
	var taikinSheets = taikinSheet.getSheets();
	Logger.log("Sheet名:" + taikinSheets[0].getSheetName());
	if ( taikinSheets[0].getSheetName() == "シート1" ) {
		var userInfo = taikinSheets[0];
	}
    userInfo.getRange(2,18).setValue("writeTaikin到達");
	var spreadsheet = file;
    userInfo.getRange(2,18).setValue("writeTaikin到達2");
    userInfo.getRange(5,17).setValue(file);
    userInfo.getRange(2,18).setValue("writeTaikin到達3");
	var sheets = spreadsheet.getSheets();
    userInfo.getRange(2,18).setValue("writeTaikin到達4");
	Logger.log("Sheet名:" + sheets[1].getSheetName());
	if ( sheets[1].getSheetName() === "原本" ) {
		var sheet = sheets[1];
    userInfo.getRange(2,18).setValue("writeTaikin到達5");
		// 日数
		Logger.log("残業した日:" + date);
		sheet.getRange(date + 4, 6).setValue(hour);
		sheet.getRange(date + 4, 7).setValue(minutes);
        userInfo.getRange(3,18).setValue(sheet.getRange(date + 4, 6).getValue());
        userInfo.getRange(4,18).setValue(sheet.getRange(date + 4, 7).getValue());

	}
	return spreadsheet;
}

/**
 * 指定月の日数を取得
 * @param  {number} year  年
 * @param  {number} month 月
 * @return {number} 指定月の日数
 */
function getLastDay(year, month) {
	return new Date(year, month, 0).getDate();
}

function createSpreadsheet(id, name) {
	// create a new file of spreadsheet (250 / day)
	var file = Drive.Files.insert({
		'title': name,
		'mimeType': 'application/vnd.google-apps.spreadsheet',
		'parents': [{'id': id}]
	});

	// open spreadsheet
	return SpreadsheetApp.openById(file.getId());
}