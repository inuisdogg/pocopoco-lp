// ===========================================================
// Google Apps Script - フォーム送信処理
// ===========================================================

// ■ 設定エリア ■

// スプレッドシートID
const SPREADSHEET_ID = "1le-REuPK_gpE0MCSEjua041W3kQsPfBcgFESarn3CeI"; 

// 履歴書保存フォルダID
const FOLDER_ID = "17eyeWRnUypn9ST1TJntPuN27j_AL5XmV"; 

// 通知先メールアドレス
const NOTIFY_EMAIL = "pocopoco.fuchu@gmail.com"; 

// ===========================================================
// メイン処理
// ===========================================================

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    console.log("=== doPost 開始 ===");
    console.log("e の存在:", !!e);
    
    // パラメータ取得の試行
    var params = {};
    var dataSource = "";
    
    // 方法1: postData.contentsから取得（JSON形式）
    if (e && e.postData && e.postData.contents) {
      console.log("postData.contents から取得を試行");
      console.log("postData.type:", e.postData.type);
      console.log("postData.contents の長さ:", e.postData.contents.length);
      
      try {
        params = JSON.parse(e.postData.contents);
        dataSource = "postData.contents";
        console.log("✓ postData.contents から取得成功");
      } catch (parseError) {
        console.error("JSONパースエラー:", parseError.toString());
        // パースに失敗した場合は次の方法を試す
      }
    }
    
    // 方法2: e.parameterから取得（フォールバック）
    if (!params.formType && e && e.parameter) {
      console.log("e.parameter から取得を試行");
      console.log("e.parameter のキー:", Object.keys(e.parameter));
      
      // e.parameterは文字列なので、必要に応じて変換
      params = e.parameter;
      dataSource = "parameter";
      console.log("✓ e.parameter から取得成功");
    }
    
    // どちらも取得できなかった場合
    if (!params.formType) {
      var errorMsg = "データを取得できませんでした。e.postData.contents と e.parameter の両方が空です。";
      console.error("エラー: " + errorMsg);
      console.error("e の内容:", JSON.stringify(e));
      return output.setContent(JSON.stringify({ 
        status: "error", 
        message: errorMsg 
      }));
    }
    
    console.log("データ取得元:", dataSource);
    console.log("formType:", params.formType);
    console.log("params のキー:", Object.keys(params));
    
    // スプレッドシートを開く
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 日時フォーマット
    var date = new Date();
    var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');

    // ===========================================================
    // ★名札（formType）による振り分け処理
    // ===========================================================
    
    // ■パターンA：利用予約フォーム
    if (params.formType === 'reservation') {
      console.log("利用予約フォームの処理を開始");
      var sheet = ss.getSheetByName('利用予約窓口'); 
      
      if (!sheet) {
        throw new Error("シート '利用予約窓口' が見つかりません");
      }
      
      // 保存
      sheet.appendRow([
        dateStr,
        params.parentName || '',
        params.childName || '',
        params.childAge || '',
        params.email || '',
        params.tel || '',
        params.certificate || '',
        params.requestType || '',
        params.message || ''
      ]);

      console.log("利用予約フォーム: スプレッドシートに保存完了");

      // メール通知
      sendNotification(
        '【相談・見学】予約が入りました',
        "利用予約フォームから申し込みがありました。\n\n" +
        "■氏名: " + (params.parentName || '') + "\n" +
        "■連絡先: " + (params.email || '') + " / " + (params.tel || '') + "\n" +
        "■内容: " + (params.requestType || '') + "\n\n" +
        "■メッセージ:\n" + (params.message || '')
      );
      
      console.log("利用予約フォーム: 処理完了");
    }

    // ■パターンB：採用フォーム
    else if (params.formType === 'recruit') {
      console.log("採用フォームの処理を開始");
      var sheet = ss.getSheetByName('採用窓口');
      
      if (!sheet) {
        throw new Error("シート '採用窓口' が見つかりません");
      }
      
      var fileUrl = "ファイル添付なし";
      var fileStatusMessage = "添付なし"; 

      // 履歴書ファイルがある場合の処理
      if (params.fileData) {
        console.log("ファイルデータが存在します。ドライブへの保存を試行");
        console.log("ファイル名:", params.fileName);
        console.log("MIMEタイプ:", params.mimeType);
        console.log("ファイルデータの長さ:", params.fileData ? params.fileData.length : 0);
        
        try {
          var contentType = params.mimeType || "application/pdf";
          var decoded = Utilities.base64Decode(params.fileData);
          var blob = Utilities.newBlob(decoded, contentType, params.fileName || "履歴書.pdf");
          
          var folder = DriveApp.getFolderById(FOLDER_ID);
          var file = folder.createFile(blob);
          
          fileUrl = file.getUrl();
          fileStatusMessage = "保存成功";
          console.log("✓ Drive保存成功: " + fileUrl);
        } catch (error) {
          fileUrl = "ドライブ保存エラー: " + error.message;
          fileStatusMessage = "保存失敗";
          console.error("✗ Drive保存失敗:", error.message);
          console.error("エラースタック:", error.stack);
        }
      } else {
        console.log("ファイルデータはありません（詳細入力応募の可能性）");
      }
      
      // シートに保存
      sheet.appendRow([
        dateStr,
        params.requestType || '',
        params.parentName || '',
        params.email || '',
        params.tel || '',
        fileUrl,
        params.message || ''
      ]);

      console.log("採用フォーム: スプレッドシートに保存完了");

      // 通知メール作成
      var subject = '【採用】応募がありました (' + fileStatusMessage + ')';
      var body = "採用フォームから応募がありました。\n\n" +
                 "■氏名: " + (params.parentName || '') + "\n" +
                 "■種別: " + (params.requestType || '') + "\n" +
                 "■連絡先: " + (params.email || '') + " / " + (params.tel || '') + "\n\n" +
                 "【履歴書ファイル】\n" + 
                 (fileStatusMessage === "保存失敗" ? 
                    "⚠️ ドライブへの保存に失敗しました。後でファイルサイズをご確認ください。\nエラーメッセージ: " + fileUrl : 
                    fileUrl) + 
                 "\n\n" +
                 "■メッセージ:\n" + (params.message || '');

      sendNotification(subject, body);
      
      console.log("採用フォーム: 処理完了");
    }

    // ■パターンC：ベビーシッター（地域投票）
    else if (params.formType === 'babysitter') {
      console.log("ベビーシッターフォームの処理を開始");
      var sheet = ss.getSheetByName('ベビーシッター利用希望窓口');
      
      if (!sheet) {
        throw new Error("シート 'ベビーシッター利用希望窓口' が見つかりません");
      }

      // 保存（日時とエリアのみ）
      sheet.appendRow([
        dateStr,
        params.area || ''
      ]);

      console.log("ベビーシッターフォーム: スプレッドシートに保存完了");

      // メール通知
      sendNotification(
        '【シッター】地域の希望投票がありました',
        "シッター利用希望地域への投票がありました。\n\n" +
        "■希望地域: " + (params.area || '') + "\n\n" +
        "※このフォームは匿名投票のため、個人情報はありません。"
      );
      
      console.log("ベビーシッターフォーム: 処理完了");
    }

    // 未対応のformTypeの場合
    else {
      var errorMsg = "未対応のフォーム種別です: " + (params.formType || 'undefined');
      console.error("エラー: " + errorMsg);
      return output.setContent(JSON.stringify({ 
        status: "error", 
        message: errorMsg 
      }));
    }

    console.log("=== doPost 正常終了 ===");
    return output.setContent(JSON.stringify({ status: "success" }));

  } catch (e) {
    console.error("=== doPost エラー発生 ===");
    console.error("エラーメッセージ: " + e.toString());
    console.error("スタックトレース: " + e.stack);
    console.error("エラーオブジェクト:", JSON.stringify(e));
    
    // エラー通知メールを送信
    try {
      MailApp.sendEmail(
        NOTIFY_EMAIL,
        "【GASエラー】フォーム送信処理でエラーが発生しました",
        "フォーム送信処理中にエラーが発生しました。\n\n" +
        "エラーメッセージ: " + e.toString() + "\n\n" +
        "スタックトレース:\n" + e.stack
      );
    } catch (mailError) {
      console.error("エラー通知メールの送信に失敗:", mailError.message);
    }
    
    return output.setContent(JSON.stringify({ 
      status: "error", 
      message: e.toString() 
    }));
  }
}

// ===========================================================
// 共通メール送信関数
// ===========================================================

function sendNotification(subject, body) {
  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    console.log("✓ メール送信成功: " + subject);
  } catch (e) {
    console.error("✗ メール送信失敗: " + e.message);
    // メール送信失敗でも処理は続行（スプレッドシートへの保存は成功しているため）
  }
}

// ===========================================================
// 権限再承認のためのテスト関数
// ===========================================================

function testDrivePermission() {
  const FOLDER_ID = "17eyeWRnUypn9ST1TJntPuN27j_AL5XmV"; 
  
  try {
    // ドライブへのアクセス権を確認（ファイルのリストを取得するだけでも権限が必要）
    DriveApp.getFolderById(FOLDER_ID).getFiles(); 
    
    // 成功した場合、権限が通っている
    MailApp.sendEmail(NOTIFY_EMAIL, "【最終テスト】ドライブアクセス成功", "GASはGoogleドライブへの書き込み権限を持っています。");
    console.log("✓ Drive権限テスト成功");
  } catch (e) {
    // 失敗した場合、権限がないかフォルダIDが間違っている
    MailApp.sendEmail(NOTIFY_EMAIL, "【最終テスト】⚠️ドライブアクセス失敗", "GASがドライブへのアクセス権を持っていません。エラー: " + e.message);
    console.error("✗ Drive権限テスト失敗:", e.message);
  }
}
