//LINE developersのメッセージ送受信設定に記載のアクセストークン
var ACCESS_TOKEN = 'Y+pvhdCTdLZqRQlYSgnh51BcNzUNFUNwdaUso5Q18wk0LoTJMAJ90O7GhOjHaIxnrrBrp6K1JYLsD41Sp91uAIZ/ZR2rAXsw6HitA71EurIPMgr0R0WDzLkX+nrlcHZ63xbkhCTSQkdwtlftLlMxIwdB04t89/1O/w1cDnyilFU=';
var ownerID = ""; // メールアドレス
var templateid = "163CopCFxrMshO8bYfqoY4Jlc3Rr9K-20_tn9njpAYwU"; // 出勤簿テンプレート
var taikinTemplate = ""; // db用シート名
var destfolderid = ""; // 保存用フォルダディレクトリid
var db = '1RUOTDUR74hNq0hYXeUAt1afvfV8w8jzjMM9sknnl0SI'; // 開発者のdb用スプレッドシート
var userName = '';
var userNo = '';

function doPost(e) {

	// WebHookで受信した応答用Token
//	var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
    var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
//	var replyToken = "";
    Logger.log(replyToken);
// ユーザーのメッセージを取得
	var userMessage = JSON.parse(e.postData.contents).events[0].message.text;
//	var userMessage = "";
    Logger.log(userMessage);
// ユーザIDを取得
	var userId = JSON.parse(e.postDate.contents).events[0].source.userId;
	userInfo.getRange(7,16).setValue(userId);
//	var userId = "";
    Logger.log(userId);
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
	userInfo.getRange(8,16).setValue(memo);

	// ログ
//	taikin.getRange(4, 4).setValue(Object.prototype.toString.call(memo));
//	taikin.getRange(4, 5).setValue(Object.prototype.toString.call(taikin.getRange(5,4).getValue()));
//	taikin.getRange(5, 6).setValue(Object.prototype.toString.call(memo.length));
//	taikin.getRange(5, 7).setValue(memo.length);

	// LINEからの初期設定受付
	// メール
	if(memo.indexOf('@gmail.com') != -1){
		userInfo.getRange(4, 16).setValue('分岐1-1');
		ownerID = memo.substring(0, memo.indexOf('@')+10);
		userName = memo.substring(11);
		var userRange = '';
		for(var i = 1;i < 5;i++){
			userRange = userInfo.getRange(i,1)
			if(userRange.isBrank()){
				userNo = i;
				userInfo.getRange(i,1).setValue(ownerID);
				userInfo.getRange(i,2).setValue(userId);
//				insertSheet(userId,i+1);
//				userInfo.getRange(3,i).setValue(i+1);
				userInfo.getRange(i,4).setValue(userName);
			}
		}
		Logger.log(userInfo.getRange(i,1).getValue);
		Logger.log(userInfo.getRange(i,2).getValue);
		Logger.log(userInfo.getRange(i,4).getValue);
		// シートで確認するログ
		userInfo.getRange(5,16).setValue(userInfo.getRange(i,1).getValue());
		userInfo.getRange(6,16).setValue(userInfo.getRange(i,2).getValue());
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
					'text':'メールアドレスと名前を設定しました。メール:' + userInfo.getRange(userNo,1).getValue() + userInfo.getRange(userNo,4).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}else if(memo.match(/フォルダ/)){
		// userNoを取得
		for(var i = 1;i < 5;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		//フォルダID
		userInfo.getRange(userNo,3).setValue(memo.substring(memo.indexOf('http')));

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
					'text':'フォルダIDを設定しました。出勤' + userInfo.getRange(userNo,3).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}
	// 定時
	else if(memo.match(/定時/)){
		// userNoを取得
		for(var i = 1;i < 5;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
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
					'text':'定時を設定しました。出勤' + userInfo.getRange(userNo,5).getValue() + '時' + userInfo.getRange(userNo,6).getValue() + '分' + '退勤' + userInfo.getRange(userNo,7).getValue() + '時' + userInfo.getRange(userNo,8).getValue() + '分',
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
	}
	// DB
//	if(memo.match(/DB用/)){
//	if(memo.match(/DB用/)){
//	db = str.substr(3);
//	} else if(memo.match(/DB用 /)){
//	db = str.substr(4);
//	}

//	destfolderid = taikin.getRange(7,7).setValue(db);

////	taikinTemplate = spreadSheetApp.create('出勤簿作成くんDB');
//	taikinTemplate = createSpreadsheet(destfolderid,'出勤簿作成くんDB');

//	UrlFetchApp.fetch(url, {
//	'headers': {
//	'Content-Type': 'application/json; charset=UTF-8',
//	'Authorization': 'Bearer ' + ACCESS_TOKEN,
//	},
//	'method': 'post',
//	'payload': JSON.stringify({
//	'replyToken': replyToken,
//	'messages': [{
//	'type': 'text',
//	'text':taikin.getRange(7,7).getValue() + 'を設定しました。',
//	},{
//	'type': 'text',
//	'text':'データベースとして使うスプレッドシートを作成しました。' + taikinTemplate.getUrl(),
//	}],
//	}),
//	});

//	}

	else if(Object.prototype.toString.call(memo) == '[object Date]'){
		// userNoを取得
		for(var i = 1;i < 5;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		// ユーザーの入力した退勤時間を設定
		userInfo.getRange(userNo,9).setValue(memo);
		// ログ
		userInfo.getRange(2, 16).setValue(taikin.getRange(9,9).getValue());
		// 退勤時間を記録する出勤簿を設定
		var file = SpreadsheetApp.openByUrl();
		// メッセージを日付と時刻に分解
		var year = memo.getFullYear();
		Logger.log("退勤（年）:" + memo.getFullYear());
		var month = memo.getMonth();
		Logger.log("退勤（月）:" + memo.getMonth());
		var date = memo.getDate();
		Logger.log("退勤（日）:" + memo.getDate());
		var hour = memo.getHours();
		Logger.log("退勤（時）:" + memo.getHours());
		var minutes = memo.getMinutes();
		if(minutes < 30){
			minutes = 0;
		} else if(minutes > 30){
			minutes = 30;
		}
		Logger.log("退勤（分）:" + minutes);

		userInfo.getRange(10,userNo).setValue(date + '日は' + hour + '時' + minutes + '分');
		var message = taikin.getRange(10,userNo).getValue();

		var spreadsheet = writeTaikin(file,date,hour,minutes);
		// ログ
		userInfo.getRange(3, 16).setValue(spreadsheet);

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
					'text':message + 'に退勤ですね。退勤時間を記録しました。'+ spreadsheet.getUrl(),
				}],
			}),
		});


		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);


	}
	else if (80 < memo.length && memo.length < 160){
		// userNoを取得
		for(var i = 1;i < 5;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		// 退勤時間を記録する出勤簿を設定
		var file = SpreadsheetApp.openByUrl(memo);
		userInfo.getRange(userNo, 11).setValue(file);
		memo  = userInfo.getRange(userNo,9).getValue();
		// メッセージを日付と時刻に分解
		var year = memo.getFullYear();
		Logger.log("退勤（年）:" + memo.getFullYear());
		var month = memo.getMonth();
		Logger.log("退勤（月）:" + memo.getMonth());
		var date = memo.getDate();
		Logger.log("退勤（日）:" + memo.getDate());
		var hour = memo.getHours();
		Logger.log("退勤（時）:" + memo.getHours());
		var minutes = memo.getMinutes();
		if(minutes < 30){
			minutes = 0;
		} else if(minutes > 30){
			minutes = 30;
		}
		Logger.log("退勤（分）:" + minutes);


		var spreadsheet = writeTaikin(file,date,hour,minutes);
		// ログ
		userInfo.getRange(3, 16).setValue(spreadsheet);

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
					'text':'退勤時間を記録しました。' + spreadsheet.getUrl(),
				}],
			}),
		});


		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}
	else if (memo.match(/出勤簿作成/)){
		// userNoを取得
		for(var i = 1;i < 5;i++){
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
		var newfile = writeSheet(month);
		// 作成した最新の出勤簿を記録
		userInfo.getRange(userNo,12).setValue(newfile.getUrl())

		// ログ
		userInfo.getRange(4, 16).setValue('分岐3-1');
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
					'text':'出勤簿を作成しました。' + userInfo.getRange(userNo,12).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	} else if (memo.match(/先月/) && memo.match(/出勤簿作成/)){
		// userNoを取得
		for(var i = 1;i < 5;i++){
			if(userInfo.getRange(i,2).getValue() == userId){
				userNo = i;
			}
		}
		// ログ
		userInfo.getRange(4, 16).setValue('分岐4');
		//日付作成
		var now = new Date();
		var month = now.getMonth();    //月
		// 出勤簿作成
		var newfile = writeSheet(month);
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
					'text':'出勤簿を作成しました。' + userInfo.getRange(userNo,12).getValue(),
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}else {
		// ログ
		userInfo.getRange(4, 16).setValue('分岐4');

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
					'text':'「退勤の日付(yyyy/mm/dd)(半角スペース)退勤時刻(hh:mm)」または「出勤簿作成」と入力してください。',
				}],
			}),
		});
		return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);

	}


}


function writeSheet(m) {

	// userNoを取得
	for(var i = 1;i < 5;i++){
		if(userInfo.getRange(i,2).getValue() == userId){
			userNo = i;
		}
	}
	// 保存フォルダを設定
	destfolderid = userInfo.getRange(i,3).getValue();
	userName = userInfo.getRange(i,4).getValue();

	//出勤簿テンプレートを指定
//	var templateid = "1kYnJn4Esd9vnoc368mjZt_8twQTooLjkzqqnEDSYU18";
	var template = DriveApp.getFileById(templateid);

	//保存フォルダを指定
//	var destfolderid = "1m7cyvz5FI-n_PuIb661N6x1yJFixIAlr";       //destination folderを略してます
	var destfolder = DriveApp.getFolderById(destfolderid);

	//日付作成
	var now = new Date();
	var year = now.getFullYear(); //年
	var month = m;    //月
	Logger.log(month);
//	var day = now.getDay(); // 曜日

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
	newfile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);

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
				sheet.getRange(j, 4).setValue(userInfo.getRange(5,userNo).getValue());
				sheet.getRange(j, 5).setValue(userInfo.getRange(6,userNo).getValue());
				sheet.getRange(j, 6).setValue(userInfo.getRange(7,userNo).getValue());
				sheet.getRange(j, 7).setValue(userInfo.getRange(8,userNo).getValue());
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
	var spreadsheet = file;
	var sheets = spreadsheet.getSheets();
	Logger.log("Sheet名:" + sheets[1].getSheetName());
	if ( sheets[1].getSheetName() === "原本" ) {
		var sheet = sheets[1];

		// 日数
		Logger.log("残業した日:" + date);
		sheet.getRange(date + 4, 6).setValue(hour);
		sheet.getRange(date + 4, 7).setValue(minutes);

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

////SpreadsheetをExcelファイルに変換してドライブに保存、Fileを返す
//function ss2xlsx(spreadsheet_id) {
//	var new_file;
//	var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?format=xlsx";
//	var options = {
//			method: "get",
//			headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
//			muteHttpExceptions: true
//	};
//	var res = UrlFetchApp.fetch(url, options);
//	if (res.getResponseCode() == 200) {
//		var ss = SpreadsheetApp.openById(spreadsheet_id);
//		new_file = DriveApp.createFile(res.getBlob()).setName(ss.getName() + ".xlsx");
//	}
//	return new_file;
//}