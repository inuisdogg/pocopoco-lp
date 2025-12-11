// ===========================================================
// Google Apps Script - フォーム送信処理
// ===========================================================

// ■ 設定エリア ■

// スプレッドシートID
const SPREADSHEET_ID = "1le-REuPK_gpE0MCSEjua041W3kQsPfBcgFESarn3CeI"; 

// 履歴書保存フォルダID
const FOLDER_ID = "17eyeWRnUypn9ST1TJntPuN27j_AL5XmV"; 

// イベント画像保存フォルダID（別フォルダに保存したい場合は、ここにフォルダIDを設定）
// フォルダIDの取得方法：Google Driveでフォルダを開く → URLの最後の部分がフォルダID
// 例：https://drive.google.com/drive/folders/【ここがフォルダID】?usp=sharing
const EVENT_IMAGE_FOLDER_ID = "1ejbctTjCIUY4kk4qrOFjtIFhNdKcRTLC"; // イベント画像専用フォルダ

// イベント一覧のスプレッドシートタブ名（新しいタブを作成する場合の名前）
const EVENT_SHEET_NAME = "イベント一覧"; // ここを変更すればタブ名を変更できます

// 通知先メールアドレス
const NOTIFY_EMAIL = "pocopoco.fuchu@gmail.com";

// 【重要】送信元メールアドレスについて
// GmailApp.sendEmailを使用することで、エイリアスを送信元として指定できます。
// info@inu.co.jpがエイリアスの場合：
// 1. そのエイリアスが設定されているGoogleアカウントでGASを実行
// 2. Gmailの設定でinfo@inu.co.jpをエイリアスとして追加（Google Workspaceの場合は管理者が設定）
// 3. このスクリプトでは自動的にinfo@inu.co.jpから送信されます
// 
// info@inu.co.jpがグループアカウントの場合：
// グループアカウントは送信元として使用できません。
// その場合は、info@inu.co.jpをエイリアスとして持つユーザーアカウントを作成するか、
// 既存のユーザーアカウントにinfo@inu.co.jpをエイリアスとして追加する必要があります。
//
// 送信元メールアドレス（エイリアスが設定されていない場合は実行アカウントのメールアドレスが使用されます）
const FROM_EMAIL = "info@inu.co.jp";
const FROM_NAME = "児童発達支援 pocopoco";

// メール署名（すべての申し込み完了メールに共通）
function getEmailSignature() {
  var signature = "\n\n";
  signature += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  signature += "株式会社INU\n";
  signature += "info@inu.co.jp\n";
  signature += "042-306-7126\n";
  signature += "ホームページ: https://inu.co.jp\n";
  signature += "\n";
  signature += "公式LINE: https://lin.ee/83hUaLY\n";
  signature += "Instagram: https://www.instagram.com/pocopoco_fuchu/\n";
  signature += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
  return signature;
}

// SNS案内テキスト（すべての申し込み完了メールに共通）
function getSNSContactInfo() {
  var info = "\n";
  info += "【ご連絡・お問い合わせについて】\n\n";
  info += "お問い合わせ、内容変更のご連絡はLINEからお願いします。\n";
  return info;
} 

// ===========================================================
// GETリクエスト処理（イベント取得用）
// ===========================================================

function doGet(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    var action = e.parameter.action;
    
    if (action === 'getEvents') {
      var isArchived = e.parameter.isArchived === 'true';
      var events = getEvents(isArchived);
      return output.setContent(JSON.stringify({ status: 'success', events: events }));
    }
    
    if (action === 'getEvent') {
      var eventId = parseInt(e.parameter.eventId);
      var event = getEventById(eventId);
      if (event) {
        return output.setContent(JSON.stringify({ status: 'success', event: event }));
      } else {
        return output.setContent(JSON.stringify({ status: 'error', message: 'イベントが見つかりません' }));
      }
    }
    
    // 状態設定アクション（GETリクエストで処理）
    if (action === 'setAlmostFull' || action === 'setFull' || action === 'reopenRegistration') {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(EVENT_SHEET_NAME);
      
      if (!sheet) {
        return output.setContent(JSON.stringify({ status: 'error', message: 'シートが見つかりません' }));
      }
      
      var eventId = parseInt(e.parameter.eventId);
      var lastRow = sheet.getLastRow();
      var found = false;
      
      for (var i = 2; i <= lastRow; i++) {
        if (sheet.getRange(i, 1).getValue() === eventId) {
          if (action === 'setAlmostFull') {
            sheet.getRange(i, 8).setValue(-1); // currentParticipantsを-1に設定
          } else if (action === 'setFull') {
            sheet.getRange(i, 8).setValue(-2); // currentParticipantsを-2に設定
          } else if (action === 'reopenRegistration') {
            sheet.getRange(i, 8).setValue(0); // currentParticipantsを0に設定
          }
          var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
          sheet.getRange(i, 12).setValue(dateStr); // 更新日時
          found = true;
          break;
        }
      }
      
      if (!found) {
        return output.setContent(JSON.stringify({ status: 'error', message: 'イベントIDが見つかりません: ' + eventId }));
      }
      
      return output.setContent(JSON.stringify({ status: 'success', message: '状態を更新しました' }));
    }
    
    return output.setContent(JSON.stringify({ status: 'error', message: '無効なアクション' }));
  } catch (error) {
    console.error('doGetエラー:', error);
    return output.setContent(JSON.stringify({ status: 'error', message: error.toString() }));
  }
}

// ===========================================================
// OPTIONSリクエスト処理（CORSプリフライト用）
// ===========================================================

function doOptions(e) {
  // CORSプリフライトリクエスト用のレスポンス
  // GASのWebアプリは自動的にCORSヘッダーを設定するため、空のレスポンスを返すだけでOK
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

// ===========================================================
// メイン処理
// ===========================================================

function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    console.log("=== 📥 doPost 開始 ===");
    console.log("e の存在:", !!e);
    console.log("e の型:", typeof e);
    
    // eがundefinedの場合のハンドリング
    // 注意: no-corsモードでリクエストを送信した場合、eがundefinedになることがあります
    // しかし、リクエストボディは送信されている可能性があるため、別の方法で取得を試みる
    if (!e) {
      console.error("❌ 警告: e が undefined です");
      console.error("⚠️ no-corsモードでリクエストを送信した場合、この警告が表示されることがあります");
      console.error("⚠️ しかし、リクエストボディは送信されている可能性があります");
      console.error("⚠️ GASのデプロイ設定を確認してください：");
      console.error("   1. デプロイの種類: ウェブアプリ");
      console.error("   2. 実行者: 自分");
      console.error("   3. アクセスできるユーザー: 全員");
      console.error("⚠️ 通常のCORSモードでリクエストを送信することを推奨します");
      
      // eがundefinedでも、リクエストボディが送信されている可能性があるため、
      // エラーを返さずに、後続の処理でパラメータ取得を試みる
      // ただし、eがundefinedの場合は、パラメータ取得が困難なため、エラーを返す
      return output.setContent(JSON.stringify({ 
        status: "error", 
        message: "リクエストデータが受信できませんでした。\n\n通常のCORSモードでリクエストを送信してください。\n\nGASのデプロイ設定:\n- 種類: ウェブアプリ\n- 実行者: 自分\n- アクセスできるユーザー: 全員" 
      }));
    }
    
    console.log("e.postData の存在:", !!(e && e.postData));
    console.log("e.parameter の存在:", !!(e && e.parameter));
    
    // e.postDataの詳細をログ出力
    if (e.postData) {
      console.log("e.postData.type:", e.postData.type);
      console.log("e.postData.contents の存在:", !!(e.postData.contents));
      console.log("e.postData.contents の長さ:", e.postData.contents ? e.postData.contents.length : 0);
      if (e.postData.contents) {
        console.log("e.postData.contents の最初の500文字:", e.postData.contents.substring(0, 500));
      }
    }
    
    // e.parameterの詳細をログ出力
    if (e.parameter) {
      console.log("e.parameter のキー:", Object.keys(e.parameter));
      console.log("e.parameter の内容:", JSON.stringify(e.parameter));
    }
    
    // パラメータ取得の試行
    var params = {};
    var dataSource = "";
    
    // 方法1: postData.contentsから取得（JSON形式）
    if (e && e.postData && e.postData.contents) {
      console.log("📥 postData.contents から取得を試行");
      console.log("postData.type:", e.postData.type);
      console.log("postData.contents の長さ:", e.postData.contents.length);
      console.log("postData.contents のサイズ:", (e.postData.contents.length / 1024 / 1024).toFixed(2), "MB");
      console.log("postData.contents の最初の200文字:", e.postData.contents.substring(0, 200));
      
      try {
        params = JSON.parse(e.postData.contents);
        dataSource = "postData.contents";
        console.log("✅ postData.contents から取得成功");
        console.log("取得したparams.action:", params.action);
        console.log("取得したparams.formType:", params.formType);
        console.log("取得したparams.imageBase64 の存在:", !!(params.imageBase64));
        console.log("取得したparams.imageFileName の存在:", !!(params.imageFileName));
        if (params.imageBase64) {
          console.log("取得したparams.imageBase64 の長さ:", params.imageBase64.length);
          console.log("取得したparams.imageBase64 の最初の50文字:", params.imageBase64.substring(0, 50));
        }
        
        // ネストされた配列が文字列になっている可能性があるため、再パースを試行
        if (params.selectedDateTimes && typeof params.selectedDateTimes === 'string') {
          try {
            params.selectedDateTimes = JSON.parse(params.selectedDateTimes);
            console.log("✓ selectedDateTimes を再パースしました");
          } catch (e) {
            console.log("selectedDateTimes の再パースに失敗:", e.message);
          }
        }
      } catch (parseError) {
        console.error("JSONパースエラー:", parseError.toString());
        // パースに失敗した場合は次の方法を試す
      }
    }
    
    // 方法2: e.parameterから取得（フォールバック）
    if (!params.formType && !params.action && e && e.parameter) {
      console.log("📥 e.parameter から取得を試行");
      console.log("e.parameter のキー:", Object.keys(e.parameter));
      
      // e.parameterは文字列なので、必要に応じて変換
      params = e.parameter;
      dataSource = "parameter";
      console.log("✅ e.parameter から取得成功");
      console.log("取得したparams.action:", params.action);
      console.log("取得したparams.formType:", params.formType);
    }
    
    // 画像アップロードの処理（FormData形式）
    if (e && e.postData && e.postData.type && e.postData.type.indexOf('multipart/form-data') !== -1) {
      console.log("FormData形式のリクエストを検出");
      var formData = e.parameter;
      if (formData.action === 'uploadImage') {
        try {
          var imageBlob = e.postData.contents;
          if (!imageBlob) {
            throw new Error('画像データがありません');
          }
          
          // Blobとして保存
          var folder = DriveApp.getFolderById(EVENT_IMAGE_FOLDER_ID);
          var fileName = 'event_' + Date.now() + '.jpg';
          var file = folder.createFile(fileName, imageBlob, 'image/jpeg');
          
          // ファイルを公開設定にする
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          
          // Google Driveの画像を直接表示するためのURL形式に変換
          var fileId = file.getId();
          var imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
          
          console.log("画像アップロード成功: " + imageUrl);
          console.log("📁 ファイルID: " + fileId);
          
          return output.setContent(JSON.stringify({ 
            status: 'success', 
            imageUrl: imageUrl 
          }));
        } catch (error) {
          console.error("画像アップロードエラー:", error);
          return output.setContent(JSON.stringify({ 
            status: 'error', 
            message: error.toString() 
          }));
        }
      }
    }
    
    // どちらも取得できなかった場合（formTypeまたはactionがない場合）
    if (!params.formType && !params.action) {
      var errorMsg = "データを取得できませんでした。e.postData.contents と e.parameter の両方が空です。";
      console.error("❌ エラー: " + errorMsg);
      console.error("e.postData の存在:", !!e.postData);
      console.error("e.postData.contents の存在:", !!(e.postData && e.postData.contents));
      console.error("e.postData.type:", e.postData ? e.postData.type : 'undefined');
      console.error("e.parameter の存在:", !!e.parameter);
      console.error("e.parameter のキー:", e.parameter ? Object.keys(e.parameter) : 'undefined');
      console.error("e の内容:", JSON.stringify(e));
      return output.setContent(JSON.stringify({ 
        status: "error", 
        message: errorMsg 
      }));
    }
    
    console.log("データ取得元:", dataSource);
    console.log("formType:", params.formType);
    console.log("action:", params.action);
    console.log("params のキー:", Object.keys(params));
    
    // デバッグ：params全体をログ出力
    console.log("=== 受信したparams全体 ===");
    console.log(JSON.stringify(params, null, 2));
    
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
      
      // 候補日時の処理
      var selectedDateTimesText = '';
      
      // デバッグログ：受信したデータを詳細に確認
      console.log("=== 候補日時データの受信確認 ===");
      console.log("params.selectedDateTimesText:", params.selectedDateTimesText);
      console.log("params.selectedDateTimesText の型:", typeof params.selectedDateTimesText);
      console.log("params.selectedDateTimes:", params.selectedDateTimes);
      console.log("params.selectedDateTimes の型:", typeof params.selectedDateTimes);
      console.log("params.selectedDateTimes が配列か:", Array.isArray(params.selectedDateTimes));
      console.log("params の全キー:", Object.keys(params));
      
      // params.selectedDateTimesが文字列の場合（JSON文字列の可能性）
      // 複数回パースが必要な場合がある（文字列化されたJSON文字列など）
      var parseAttempts = 0;
      var maxParseAttempts = 3;
      var currentValue = params.selectedDateTimes;
      
      while (typeof currentValue === 'string' && parseAttempts < maxParseAttempts) {
        try {
          var parsed = JSON.parse(currentValue);
          if (Array.isArray(parsed)) {
            params.selectedDateTimes = parsed;
            console.log("✓ selectedDateTimes を" + (parseAttempts + 1) + "回目のパースで配列として取得しました");
            break;
          } else if (typeof parsed === 'string') {
            // まだ文字列の場合は再パースを試行
            currentValue = parsed;
            parseAttempts++;
            console.log("再パースが必要です（" + parseAttempts + "回目）");
          } else {
            console.log("パース結果が配列でも文字列でもありません:", typeof parsed);
            break;
          }
        } catch (e) {
          console.log("パースに失敗（" + (parseAttempts + 1) + "回目）:", e.message);
          // パースできない場合は文字列のまま処理を試みる
          break;
        }
      }
      
      // 最終的に文字列のままの場合、手動で配列に変換を試みる
      if (typeof params.selectedDateTimes === 'string' && params.selectedDateTimes.trim() !== '') {
        console.log("selectedDateTimes が文字列のままです。手動パースを試行します");
        console.log("文字列の内容:", params.selectedDateTimes.substring(0, 200)); // 最初の200文字を表示
      }
      
      // 候補日時の取得を試行（複数の方法で）
      if (params.selectedDateTimesText && params.selectedDateTimesText.trim() !== '') {
        // 方法1: テキスト形式の候補日時がある場合（最優先）
        selectedDateTimesText = params.selectedDateTimesText;
        console.log("✓ 方法1: selectedDateTimesText から取得:", selectedDateTimesText);
      } else if (params.selectedDateTimes) {
        // 方法2: 配列形式の候補日時がある場合
        if (Array.isArray(params.selectedDateTimes)) {
          console.log("配列の要素数:", params.selectedDateTimes.length);
          if (params.selectedDateTimes.length > 0) {
            console.log("配列の最初の要素:", JSON.stringify(params.selectedDateTimes[0]));
          }
          
          selectedDateTimesText = params.selectedDateTimes.map(function(dt, index) {
            console.log("要素" + index + ":", JSON.stringify(dt));
            
            // displayTextを最優先
            if (dt && dt.displayText) {
              console.log("  → displayTextを使用:", dt.displayText);
              return dt.displayText;
            } 
            // dateLabelとtimeLabelから作成
            else if (dt && dt.dateLabel && dt.timeLabel) {
              var result = dt.dateLabel + ' ' + dt.timeLabel;
              console.log("  → dateLabel + timeLabelから作成:", result);
              return result;
            } 
            // dateとtimeから作成
            else if (dt && dt.date && dt.time) {
              try {
                var dateObj = new Date(dt.date);
                if (!isNaN(dateObj.getTime())) {
                  var weekdays = ['日', '月', '火', '水', '木', '金', '土'];
                  var dateLabel = (dateObj.getMonth() + 1) + '月' + dateObj.getDate() + '日（' + weekdays[dateObj.getDay()] + '）';
                  var timeLabel = dt.time + '-' + (dt.end || dt.time);
                  var result = dateLabel + ' ' + timeLabel;
                  console.log("  → date + timeから作成:", result);
                  return result;
                }
              } catch (e) {
                console.log("  → 日付パースエラー:", e.message);
              }
            }
            console.log("  → データが不完全です");
            return '';
          }).filter(function(text) {
            return text && text.trim() !== '';
          }).join('、');
          
          console.log("✓ 方法2: selectedDateTimes 配列から取得:", selectedDateTimesText);
        } else {
          console.log("⚠ selectedDateTimes が配列ではありません。型:", typeof params.selectedDateTimes);
          console.log("値:", params.selectedDateTimes);
        }
      } else {
        console.log("⚠ 候補日時データが見つかりません");
        console.log("params.selectedDateTimesText の存在:", !!params.selectedDateTimesText);
        console.log("params.selectedDateTimes の存在:", !!params.selectedDateTimes);
      }
      
      console.log("最終的な候補日時データ:", selectedDateTimesText);
      console.log("最終的な候補日時データの長さ:", selectedDateTimesText ? selectedDateTimesText.length : 0);

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
        selectedDateTimesText || '',  // 候補日時を追加
        params.message || ''
      ]);

      console.log("利用予約フォーム: スプレッドシートに保存完了");

      // メール通知
      var emailBody = "利用予約フォームから申し込みがありました。\n\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "【基本情報】\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "■氏名: " + (params.parentName || '') + "\n" +
        "■お子さまのお名前: " + (params.childName || '') + "\n" +
        "■お子さまの年齢: " + (params.childAge || '') + "\n" +
        "■連絡先: " + (params.email || '') + " / " + (params.tel || '') + "\n" +
        "■受給者証の有無: " + (params.certificate || '') + "\n" +
        "■お問い合わせ内容: " + (params.requestType || '') + "\n\n";
      
      // 見学・面談希望の場合は候補日時を強調表示
      var isVisitOrConsultation = params.requestType && 
        (params.requestType.includes('見学希望') || params.requestType.includes('面談希望'));
      
      if (isVisitOrConsultation) {
        emailBody += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
          "【希望候補日時】\n" +
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
        
        if (selectedDateTimesText && selectedDateTimesText.trim() !== '') {
          // 候補日時を番号付きで表示
          var dateTimesArray = selectedDateTimesText.split('、');
          dateTimesArray.forEach(function(dt, index) {
            emailBody += "候補" + (index + 1) + ": " + dt + "\n";
          });
        } else {
          emailBody += "⚠️ 候補日時が選択されていません\n";
        }
        emailBody += "\n";
      } else if (selectedDateTimesText && selectedDateTimesText.trim() !== '') {
        // 見学・面談希望以外でも候補日時がある場合は表示
        emailBody += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
          "【希望候補日時】\n" +
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
        var dateTimesArray = selectedDateTimesText.split('、');
        dateTimesArray.forEach(function(dt, index) {
          emailBody += "候補" + (index + 1) + ": " + dt + "\n";
        });
        emailBody += "\n";
      }
      
      // メッセージがある場合のみ表示
      if (params.message && params.message.trim() !== '') {
        emailBody += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
          "【その他ご質問など】\n" +
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
          (params.message || '') + "\n\n";
      }
      
      emailBody += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
      
      sendNotification(
        '【相談・見学】予約が入りました',
        emailBody
      );
      
      // 申し込み完了メールを申し込み者に送信
      if (params.email && params.email.trim() !== '') {
        try {
          var userEmailBody = "この度は、ご相談・見学のお申し込みありがとうございます。\n\n";
          userEmailBody += "【お申し込み内容】\n";
          userEmailBody += "お問い合わせ内容: " + (params.requestType || '') + "\n";
          
          if (selectedDateTimesText && selectedDateTimesText.trim() !== '') {
            userEmailBody += "\n希望候補日時:\n";
            var dateTimesArray = selectedDateTimesText.split('、');
            dateTimesArray.forEach(function(dt, index) {
              userEmailBody += "候補" + (index + 1) + ": " + dt + "\n";
            });
          }
          
          userEmailBody += "\n【お申し込み者情報】\n";
          userEmailBody += "保護者様のお名前: " + (params.parentName || '') + "\n";
          if (params.childName) {
            userEmailBody += "お子さまのお名前: " + params.childName + "\n";
          }
          if (params.childAge) {
            userEmailBody += "お子さまの年齢: " + params.childAge + "\n";
          }
          if (params.certificate) {
            userEmailBody += "受給者証の有無: " + params.certificate + "\n";
          }
          
          if (params.message && params.message.trim() !== '') {
            userEmailBody += "\n【ご質問・ご要望】\n";
            userEmailBody += params.message + "\n";
          }
          
          userEmailBody += getSNSContactInfo();
          userEmailBody += getEmailSignature();
          
          sendEmailWithFrom(
            params.email,
            '【pocopoco】ご相談・見学のお申し込みありがとうございます',
            userEmailBody
          );
          
          console.log("✅ 申し込み完了メールを送信しました: " + params.email);
        } catch (emailError) {
          console.error("❌ 申し込み完了メールの送信に失敗しました:", emailError);
        }
      } else {
        console.log("⚠️ メールアドレスが入力されていないため、申し込み完了メールを送信しませんでした");
      }
      
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
      
      // 申し込み完了メールを応募者に送信
      if (params.email && params.email.trim() !== '') {
        try {
          var userEmailBody = "この度は、採用応募のお申し込みありがとうございます。\n\n";
          userEmailBody += "【応募内容】\n";
          userEmailBody += "応募方法: " + (params.requestType || '') + "\n";
          
          if (fileStatusMessage !== "添付なし" && fileUrl && fileUrl !== "ファイル添付なし") {
            userEmailBody += "履歴書: " + (fileStatusMessage === "保存成功" ? "正常に受信いたしました" : "受信に問題がありました") + "\n";
          }
          
          userEmailBody += "\n【応募者情報】\n";
          userEmailBody += "お名前: " + (params.parentName || '') + "\n";
          
          if (params.message && params.message.trim() !== '') {
            userEmailBody += "\n【メッセージ】\n";
            userEmailBody += params.message + "\n";
          }
          
          userEmailBody += getSNSContactInfo();
          userEmailBody += getEmailSignature();
          
          sendEmailWithFrom(
            params.email,
            '【pocopoco】採用応募のお申し込みありがとうございます',
            userEmailBody
          );
          
          console.log("✅ 申し込み完了メールを送信しました: " + params.email);
        } catch (emailError) {
          console.error("❌ 申し込み完了メールの送信に失敗しました:", emailError);
        }
      } else {
        console.log("⚠️ メールアドレスが入力されていないため、申し込み完了メールを送信しませんでした");
      }
      
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

    // ■パターンD：イベント申し込み
    else if (params.formType === 'event') {
      console.log("イベント申し込みフォームの処理を開始");
      var sheet = ss.getSheetByName('イベント申し込み');
      
      if (!sheet) {
        // シートが存在しない場合は作成
        sheet = ss.insertSheet('イベント申し込み');
        // ヘッダー行を追加
        sheet.appendRow([
          '送信日時',
          'イベントID',
          'イベント名',
          'イベント日',
          'イベント時間',
          '保護者名',
          'お子さま名',
          'お子さま年齢',
          'メールアドレス',
          '電話番号',
          '参加人数',
          '発達の悩み',
          '発達の悩み詳細',
          'その他メッセージ'
        ]);
        // ヘッダー行を太字にする
        var headerRange = sheet.getRange(1, 1, 1, 14);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#E8E8E8');
      } else {
        // 既存のシートに列が不足している場合は追加
        var lastColumn = sheet.getLastColumn();
        if (lastColumn < 14) {
          // ヘッダー行を確認して、不足している列を追加
          if (lastColumn < 12) {
            sheet.getRange(1, 12).setValue('発達の悩み');
          }
          if (lastColumn < 13) {
            sheet.getRange(1, 13).setValue('発達の悩み詳細');
          }
          if (lastColumn < 14) {
            sheet.getRange(1, 14).setValue('その他メッセージ');
          }
          // ヘッダー行を太字にする
          var headerRange = sheet.getRange(1, 1, 1, 14);
          headerRange.setFontWeight('bold');
          headerRange.setBackground('#E8E8E8');
        }
      }
      
      // 保存
      sheet.appendRow([
        dateStr,
        params.eventId || '',
        params.eventTitle || '',
        params.eventDate || '',
        params.eventTime || '',
        params.parentName || '',
        params.childName || '',
        params.childAge || '',
        params.email || '',
        params.tel || '',
        params.participants || '',
        params.developmentConcern || '',
        params.developmentConcernDetail || '',
        params.message || ''
      ]);

      console.log("イベント申し込みフォーム: スプレッドシートに保存完了");

      // メール通知
      var emailBody = "イベント申し込みフォームから申し込みがありました。\n\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "【イベント情報】\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "■イベント名: " + (params.eventTitle || '') + "\n" +
        "■開催日: " + (params.eventDate || '') + "\n" +
        "■開催時間: " + (params.eventTime || '') + "\n\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "【申し込み者情報】\n" +
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
        "■保護者様のお名前: " + (params.parentName || '') + "\n" +
        "■お子さまのお名前: " + (params.childName || '') + "\n" +
        "■お子さまの年齢: " + (params.childAge || '') + "\n" +
        "■連絡先: " + (params.email || '') + " / " + (params.tel || '') + "\n" +
        "■参加人数: " + (params.participants || '') + "\n" +
        "■発達の悩み: " + (params.developmentConcern || '') + "\n";
      
      // 発達の悩み詳細がある場合のみ表示
      if (params.developmentConcernDetail && params.developmentConcernDetail.trim() !== '') {
        emailBody += "■発達の悩み詳細: " + params.developmentConcernDetail + "\n";
      }
      
      emailBody += "\n";
      
      // メッセージがある場合のみ表示
      if (params.message && params.message.trim() !== '') {
        emailBody += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
          "【その他ご質問など】\n" +
          "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
          (params.message || '') + "\n\n";
      }
      
      emailBody += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
      
      sendNotification(
        '【イベント申し込み】' + (params.eventTitle || 'イベント') + 'への申し込みがありました',
        emailBody
      );
      
      // 申し込み完了メールを申し込み者に送信
      if (params.email && params.email.trim() !== '') {
        try {
          // イベント情報を取得
          var eventId = parseInt(params.eventId);
          var event = getEventById(eventId);
          
          var userEmailBody = "この度は、イベントへのお申し込みありがとうございます。\n\n";
          userEmailBody += "【お申し込み内容】\n";
          userEmailBody += "イベント名: " + (params.eventTitle || '') + "\n";
          userEmailBody += "開催日: " + (params.eventDate || '') + "\n";
          userEmailBody += "開催時間: " + (params.eventTime || '') + "\n";
          
          if (event) {
            // 会場はvenueフィールドから取得（registrantではない）
            if (event.venue && event.venue.trim() !== '') {
              userEmailBody += "会場: " + event.venue + "\n";
            }
            if (event.fee && event.fee.trim() !== '') {
              userEmailBody += "参加費: " + event.fee + "\n";
            }
            if (event.items && event.items.trim() !== '') {
              userEmailBody += "持ち物: " + event.items + "\n";
            }
          }
          
          userEmailBody += "\n【お申し込み者情報】\n";
          userEmailBody += "保護者様のお名前: " + (params.parentName || '') + "\n";
          userEmailBody += "お子さまのお名前: " + (params.childName || '') + "\n";
          if (params.childAge) {
            userEmailBody += "お子さまの年齢: " + params.childAge + "\n";
          }
          userEmailBody += "参加人数: " + (params.participants || '') + "\n";
          
          if (params.message && params.message.trim() !== '') {
            userEmailBody += "\n【ご質問・ご要望】\n";
            userEmailBody += params.message + "\n";
          }
          
          userEmailBody += getSNSContactInfo();
          userEmailBody += getEmailSignature();
          
          sendEmailWithFrom(
            params.email,
            '【pocopoco】イベント申し込み完了のお知らせ',
            userEmailBody
          );
          
          console.log("✅ 申し込み完了メールを送信しました: " + params.email);
        } catch (emailError) {
          console.error("❌ 申し込み完了メールの送信に失敗しました:", emailError);
          // メール送信エラーは処理を続行（スプレッドシートへの保存は成功しているため）
        }
      } else {
        console.log("⚠️ メールアドレスが入力されていないため、申し込み完了メールを送信しませんでした");
      }
      
      console.log("イベント申し込みフォーム: 処理完了");
    }

    // ■パターンE：イベント管理（作成・更新・削除・アーカイブ・状態設定）
    else if (params.action === 'createEvent' || params.action === 'updateEvent' || params.action === 'deleteEvent' || params.action === 'archiveEvent' || params.action === 'setAlmostFull' || params.action === 'setFull' || params.action === 'reopenRegistration') {
      console.log("=== ✅ イベント管理処理を開始 ===");
      console.log("action: " + params.action);
      console.log("params のキー: " + Object.keys(params));
      console.log("params: " + JSON.stringify(params));
      console.log("EVENT_SHEET_NAME: " + EVENT_SHEET_NAME);
      console.log("EVENT_IMAGE_FOLDER_ID: " + EVENT_IMAGE_FOLDER_ID);
      console.log("SPREADSHEET_ID: " + SPREADSHEET_ID);
      
      try {
        // スプレッドシートは既に開かれている（doPostの最初で開いている）
        console.log("スプレッドシートID: " + SPREADSHEET_ID);
        
        var sheet = ss.getSheetByName(EVENT_SHEET_NAME);
        console.log("📊 シート取得結果: " + (sheet ? "✅ 成功" : "❌ 失敗"));
        console.log("📊 シート名: " + EVENT_SHEET_NAME);
        console.log("📊 スプレッドシートの全シート名:", ss.getSheets().map(function(s) { return s.getName(); }));
      
      if (!sheet) {
        console.log("📊 シートが存在しないため、新規作成します");
        // シートが存在しない場合は作成
        sheet = ss.insertSheet(EVENT_SHEET_NAME);
        console.log("✅ シートを作成しました: " + EVENT_SHEET_NAME);
        // ヘッダー行を追加
        sheet.appendRow([
          'ID',                    // 1列目
          'タイトル',              // 2列目
          'カテゴリ',              // 3列目
          '説明',                  // 4列目
          '開催日',                // 5列目
          '開催時間',              // 6列目
          '定員',                  // 7列目
          '現在の参加者数',        // 8列目
          '画像URL',               // 9列目: 画像Base64データ（Google Drive不要）
          'アーカイブ',            // 10列目
          '作成日時',              // 11列目
          '更新日時',              // 12列目
          '掲載期限',              // 13列目
          '会場',                  // 14列目
          '参加費',                // 15列目
          '持ち物',                // 16列目
          '登録者'                 // 17列目
        ]);
        // ヘッダー行を太字にする
        var headerRange = sheet.getRange(1, 1, 1, 17);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#E8E8E8');
        console.log("✅ ヘッダー行を追加しました");
      } else {
        console.log("✅ 既存のシートを使用します");
        console.log("📊 シートの最終行: " + sheet.getLastRow());
        console.log("📊 シートの列数: " + sheet.getLastColumn());
        
        // 既存のシートに列が不足している場合は追加
        if (sheet.getLastColumn() < 17) {
          console.log("⚠️ シートの列数が不足しています（" + sheet.getLastColumn() + "列）。17列に拡張します");
          // 列を追加（既存の列はそのまま）
          if (sheet.getLastColumn() < 13) {
            sheet.getRange(1, 13).setValue('掲載期限');
          }
          if (sheet.getLastColumn() < 14) {
            sheet.getRange(1, 14).setValue('会場');
          }
          if (sheet.getLastColumn() < 15) {
            sheet.getRange(1, 15).setValue('参加費');
          }
          if (sheet.getLastColumn() < 16) {
            sheet.getRange(1, 16).setValue('持ち物');
          }
          if (sheet.getLastColumn() < 17) {
            sheet.getRange(1, 17).setValue('登録者');
          }
          // ヘッダー行を太字にする
          var headerRange = sheet.getRange(1, 1, 1, 17);
          headerRange.setFontWeight('bold');
          headerRange.setBackground('#E8E8E8');
          console.log("✅ シートの列数を17列に拡張しました");
        }
      }
      
      var date = new Date();
      var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      
      // 画像をBase64データとして直接スプレッドシートに保存（Google Drive不要）
      var imageBase64Data = '';
      console.log("🖼️ ========== 画像処理開始 ==========");
      console.log("🖼️ params.imageBase64 の存在: " + !!(params.imageBase64));
      console.log("🖼️ params.imageBase64 の型: " + typeof params.imageBase64);
      
      if (params.imageBase64) {
        console.log("🖼️ params.imageBase64 の長さ: " + params.imageBase64.length);
        console.log("🖼️ params.imageBase64 の最初の50文字: " + params.imageBase64.substring(0, 50));
        
        // Base64データをそのまま保存（Google Driveへのアップロード不要）
        imageBase64Data = params.imageBase64;
        
        // スプレッドシートのセルサイズ制限（50,000文字）をチェック
        if (imageBase64Data.length > 50000) {
          console.warn("⚠️ 警告: Base64データが大きすぎます（" + imageBase64Data.length + "文字）");
          console.warn("⚠️ スプレッドシートのセルサイズ制限（50,000文字）を超える可能性があります");
          console.warn("⚠️ 画像を圧縮するか、サイズを小さくすることを推奨します");
        }
        
        console.log("✅ Base64データをそのまま使用（Google Drive不要）");
        console.log("🖼️ 保存するBase64データの長さ: " + imageBase64Data.length + "文字");
      } else {
        console.log("ℹ️ 画像データなし（スキップ）");
      }
      
      console.log("🖼️ ========== 最終的な画像データ ==========");
      console.log("🖼️ imageBase64Data: " + (imageBase64Data ? imageBase64Data.substring(0, 50) + '... (長さ: ' + imageBase64Data.length + ')' : '(空)'));
      
      // イベント作成
      if (params.action === 'createEvent') {
        console.log("📝 イベント作成処理を開始");
        console.log("📝 params.title: " + (params.title || '(空)'));
        console.log("📝 params.category: " + (params.category || '(空)'));
        console.log("📝 params.date: " + (params.date || '(空)'));
        console.log("📝 params.time: " + (params.time || '(空)'));
        console.log("📝 params.capacity: " + (params.capacity || '(空)'));
        console.log("📝 imageUrl: " + (imageUrl || '(空)'));
        
        try {
          // 最大IDを取得
          var lastRow = sheet.getLastRow();
          console.log("📝 現在の最終行: " + lastRow);
          var maxId = 0;
          if (lastRow > 1) {
            var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
            console.log("📝 既存のID数: " + ids.length);
            for (var i = 0; i < ids.length; i++) {
              if (ids[i][0] > maxId) {
                maxId = ids[i][0];
              }
            }
          }
          var newId = maxId + 1;
          console.log("📝 新しいID: " + newId);
          
          var rowData = [
            newId,                                    // 1列目: ID
            params.title || '',                       // 2列目: タイトル
            params.category || '',                    // 3列目: カテゴリ
            params.description || '',                  // 4列目: 説明
            params.date || '',                        // 5列目: 開催日
            params.time || '',                        // 6列目: 開催時間
            params.capacity || 0,                     // 7列目: 定員
            params.currentParticipants || 0,         // 8列目: 現在の参加者数
            imageBase64Data || '',                    // 9列目: 画像Base64データ（Google Drive不要）
            'false',                                  // 10列目: アーカイブ
            dateStr,                                  // 11列目: 作成日時
            dateStr,                                  // 12列目: 更新日時
            params.displayDeadline || '',             // 13列目: 掲載期限
            params.venue || 'pocopoco 東京都府中新町1丁目71-7', // 14列目: 会場
            params.fee || '無料',                     // 15列目: 参加費
            params.items || '',                       // 16列目: 持ち物
            params.registrant || ''                   // 17列目: 登録者
          ];
          
          console.log("📝 ========== 行データ作成完了 ==========");
          console.log("📝 新しいID: " + newId);
          console.log("📝 タイトル: " + (params.title || '(空)'));
          console.log("📝 カテゴリ: " + (params.category || '(空)'));
          console.log("📝 開催日: " + (params.date || '(空)'));
          console.log("📝 開催時間: " + (params.time || '(空)'));
          console.log("📝 定員: " + (params.capacity || 0));
          console.log("📝 現在の参加者数: " + (params.currentParticipants || 0));
          console.log("📝 画像Base64データ（9列目）: " + (imageBase64Data ? imageBase64Data.substring(0, 50) + '... (長さ: ' + imageBase64Data.length + ')' : '(空)'));
          console.log("📝 画像Base64データの長さ: " + (imageBase64Data ? imageBase64Data.length : 0));
          console.log("📝 アーカイブ: false");
          console.log("📝 作成日時: " + dateStr);
          console.log("📝 更新日時: " + dateStr);
          console.log("📝 追加する行データ（JSON）: " + JSON.stringify(rowData));
          console.log("📝 追加前の最終行: " + sheet.getLastRow());
          
          // スプレッドシートに追加
          console.log("📝 appendRow実行開始");
          sheet.appendRow(rowData);
          console.log("✅ appendRow実行完了");
          
          var addedRowNumber = sheet.getLastRow();
          console.log("📝 追加後の最終行: " + addedRowNumber);
          
          // 追加した行の内容を確認（特に9列目の画像URL）
          console.log("📝 ========== 追加した行の内容確認 ==========");
          var addedRow = sheet.getRange(addedRowNumber, 1, 1, 17).getValues()[0];
          console.log("📝 行番号: " + addedRowNumber);
          console.log("📝 1列目（ID）: " + addedRow[0]);
          console.log("📝 2列目（タイトル）: " + addedRow[1]);
          console.log("📝 3列目（カテゴリ）: " + addedRow[2]);
          console.log("📝 4列目（説明）: " + (addedRow[3] ? addedRow[3].substring(0, 50) + '...' : '(空)'));
          console.log("📝 5列目（開催日）: " + addedRow[4]);
          console.log("📝 6列目（開催時間）: " + addedRow[5]);
          console.log("📝 7列目（定員）: " + addedRow[6]);
          console.log("📝 8列目（現在の参加者数）: " + addedRow[7]);
          console.log("📝 9列目（画像Base64データ）: " + (addedRow[8] ? addedRow[8].substring(0, 50) + '... (長さ: ' + addedRow[8].length + ')' : '(空)'));
          console.log("📝 9列目の長さ: " + (addedRow[8] ? addedRow[8].length : 0));
          console.log("📝 10列目（アーカイブ）: " + addedRow[9]);
          console.log("📝 11列目（作成日時）: " + addedRow[10]);
          console.log("📝 12列目（更新日時）: " + addedRow[11]);
          console.log("📝 13列目（掲載期限）: " + addedRow[12]);
          console.log("📝 14列目（会場）: " + addedRow[13]);
          console.log("📝 15列目（参加費）: " + addedRow[14]);
          console.log("📝 16列目（持ち物）: " + addedRow[15]);
          console.log("📝 17列目（登録者）: " + addedRow[16]);
          console.log("📝 追加した行の内容（JSON）: " + JSON.stringify(addedRow));
          
          // 9列目を直接確認
          var imageBase64Cell = sheet.getRange(addedRowNumber, 9).getValue();
          console.log("📝 ========== 9列目の直接確認 ==========");
          console.log("📝 9列目の値: " + (imageBase64Cell ? imageBase64Cell.substring(0, 50) + '... (長さ: ' + imageBase64Cell.length + ')' : '(空)'));
          console.log("📝 9列目の値の型: " + typeof imageBase64Cell);
          console.log("📝 9列目の値の長さ: " + (imageBase64Cell ? imageBase64Cell.length : 0));
          
          // 念のため、もう一度確認
          var verifyRow = sheet.getRange(addedRowNumber, 1, 1, 17).getValues()[0];
          console.log("📝 ========== 検証: 最終行の内容 ==========");
          console.log("📝 検証: 9列目（画像Base64データ）: " + (verifyRow[8] ? verifyRow[8].substring(0, 50) + '... (長さ: ' + verifyRow[8].length + ')' : '(空)'));
          console.log("📝 検証: 9列目の長さ: " + (verifyRow[8] ? verifyRow[8].length : 0));
          
          // もし9列目が空の場合は、直接setValueで設定を試みる
          if (!imageBase64Cell || imageBase64Cell === '') {
            console.log("⚠️ 9列目が空のため、直接setValueで設定を試みます");
            console.log("⚠️ 設定する画像Base64データ: " + (imageBase64Data ? imageBase64Data.substring(0, 50) + '... (長さ: ' + imageBase64Data.length + ')' : '(空)'));
            sheet.getRange(addedRowNumber, 9).setValue(imageBase64Data);
            var retryImageBase64 = sheet.getRange(addedRowNumber, 9).getValue();
            console.log("⚠️ 再設定後の9列目の値: " + (retryImageBase64 ? retryImageBase64.substring(0, 50) + '... (長さ: ' + retryImageBase64.length + ')' : '(空)'));
          }
          
          console.log("✅ イベント作成完了: ID=" + newId);
          
          // イベントIDを返す
          return output.setContent(JSON.stringify({ 
            status: 'success', 
            message: 'イベントを作成しました',
            eventId: newId
          }));
        } catch (saveError) {
          console.error("❌ スプレッドシートへの保存エラー:", saveError);
          console.error("❌ エラーメッセージ:", saveError.toString());
          console.error("❌ エラースタック:", saveError.stack);
          throw saveError; // エラーを再スローして、外側のtry-catchでキャッチされるようにする
        }
      }
      // イベント更新
      else if (params.action === 'updateEvent') {
        var eventId = parseInt(params.eventId);
        var lastRow = sheet.getLastRow();
        var found = false;
        
        for (var i = 2; i <= lastRow; i++) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            sheet.getRange(i, 2).setValue(params.title || '');
            sheet.getRange(i, 3).setValue(params.category || '');
            sheet.getRange(i, 4).setValue(params.description || '');
            sheet.getRange(i, 5).setValue(params.date || '');
            sheet.getRange(i, 6).setValue(params.time || '');
            sheet.getRange(i, 7).setValue(params.capacity || 0);
            sheet.getRange(i, 8).setValue(params.currentParticipants || 0);
            // 新しい画像がある場合は更新、ない場合は既存の画像Base64データを保持
            if (imageBase64Data) {
              sheet.getRange(i, 9).setValue(imageBase64Data);
            }
            sheet.getRange(i, 12).setValue(dateStr);
            sheet.getRange(i, 13).setValue(params.displayDeadline || '');
            sheet.getRange(i, 14).setValue(params.venue || 'pocopoco 東京都府中新町1丁目71-7');
            sheet.getRange(i, 15).setValue(params.fee || '無料');
            sheet.getRange(i, 16).setValue(params.items || '');
            sheet.getRange(i, 17).setValue(params.registrant || '');
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        console.log("イベント更新完了: ID=" + eventId);
      }
      
      // イベント削除
      else if (params.action === 'deleteEvent') {
        var eventId = parseInt(params.eventId);
        var lastRow = sheet.getLastRow();
        var found = false;
        
        for (var i = lastRow; i >= 2; i--) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            sheet.deleteRow(i);
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        console.log("イベント削除完了: ID=" + eventId);
      }
      
      // イベントアーカイブ
      else if (params.action === 'archiveEvent') {
        var eventId = parseInt(params.eventId);
        var lastRow = sheet.getLastRow();
        var found = false;
        
        for (var i = 2; i <= lastRow; i++) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            sheet.getRange(i, 10).setValue('true');
            sheet.getRange(i, 12).setValue(dateStr);
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        console.log("イベントアーカイブ完了: ID=" + eventId);
      }
      
      // 残りわずかに設定（currentParticipantsを-1に設定）
      else if (params.action === 'setAlmostFull') {
        var eventId = parseInt(params.eventId);
        var lastRow = sheet.getLastRow();
        var found = false;
        
        for (var i = 2; i <= lastRow; i++) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            sheet.getRange(i, 8).setValue(-1); // currentParticipantsを-1に設定
            sheet.getRange(i, 12).setValue(dateStr); // 更新日時
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        console.log("残りわずか設定完了: ID=" + eventId);
      }
      
      // 満員に設定（currentParticipantsを-2に設定）
      else if (params.action === 'setFull') {
        var eventId = parseInt(params.eventId);
        var lastRow = sheet.getLastRow();
        var found = false;
        
        for (var i = 2; i <= lastRow; i++) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            sheet.getRange(i, 8).setValue(-2); // currentParticipantsを-2に設定
            sheet.getRange(i, 12).setValue(dateStr); // 更新日時
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        console.log("満員設定完了: ID=" + eventId);
      }
      
      // 受付再開（currentParticipantsを0に設定）
      else if (params.action === 'reopenRegistration') {
        var eventId = parseInt(params.eventId);
        var lastRow = sheet.getLastRow();
        var found = false;
        
        for (var i = 2; i <= lastRow; i++) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            sheet.getRange(i, 8).setValue(0); // currentParticipantsを0に設定
            sheet.getRange(i, 12).setValue(dateStr); // 更新日時
            found = true;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        console.log("受付再開完了: ID=" + eventId);
      }
      
      console.log("=== ✅ イベント管理処理完了 ===");
      return output.setContent(JSON.stringify({ 
        status: "success", 
        message: "イベントの処理が完了しました" 
      }));
      } catch (error) {
        console.error("❌ イベント管理処理エラー:", error);
        console.error("エラーメッセージ:", error.toString());
        console.error("エラースタック:", error.stack);
        // エラー通知メールを送信
        try {
          sendEmailWithFrom(
            NOTIFY_EMAIL,
            "【GASエラー】イベント管理処理でエラーが発生しました",
            "イベント管理処理中にエラーが発生しました。\n\n" +
            "action: " + (params.action || 'undefined') + "\n" +
            "エラーメッセージ: " + error.toString() + "\n\n" +
            "スタックトレース:\n" + error.stack
          );
        } catch (mailError) {
          console.error("エラー通知メールの送信に失敗:", mailError.message);
        }
        return output.setContent(JSON.stringify({ 
          status: "error", 
          message: error.toString() 
        }));
      }
    }

    // ■パターンF：イベント画像アップロード（別途アップロード）
    else if (params.action === 'uploadEventImage') {
      console.log("=== ✅ イベント画像アップロード処理を開始 ===");
      console.log("eventId: " + (params.eventId || '(空)'));
      console.log("imageBase64 の存在: " + !!(params.imageBase64));
      console.log("imageFileName: " + (params.imageFileName || '(空)'));
      
      try {
        var eventId = parseInt(params.eventId);
        if (!eventId) {
          throw new Error('イベントIDが指定されていません');
        }
        
        var sheet = ss.getSheetByName(EVENT_SHEET_NAME);
        if (!sheet) {
          throw new Error('イベント一覧シートが見つかりません');
        }
        
        // イベントを検索
        var lastRow = sheet.getLastRow();
        var found = false;
        var eventRow = 0;
        
        for (var i = 2; i <= lastRow; i++) {
          if (sheet.getRange(i, 1).getValue() === eventId) {
            found = true;
            eventRow = i;
            break;
          }
        }
        
        if (!found) {
          throw new Error('イベントIDが見つかりません: ' + eventId);
        }
        
        // 画像をアップロード
        var imageUrl = '';
        if (params.imageBase64) {
          console.log("🖼️ 画像アップロード処理を開始");
          var decoded = Utilities.base64Decode(params.imageBase64);
          var fileName = params.imageFileName || 'event_' + eventId + '_' + Date.now() + '.jpg';
          var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);
          
          var folder = DriveApp.getFolderById(EVENT_IMAGE_FOLDER_ID);
          var file = folder.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          
          var fileId = file.getId();
          imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
          
          console.log("✅ 画像アップロード成功");
          console.log("📁 ファイルID: " + fileId);
          console.log("🖼️ 画像URL: " + imageUrl);
          
          // スプレッドシートの画像URLを更新（9列目）
          console.log("📝 ========== スプレッドシートの画像URL更新開始 ==========");
          console.log("📝 イベント行番号: " + eventRow);
          console.log("📝 更新前の9列目の値: " + sheet.getRange(eventRow, 9).getValue());
          console.log("📝 設定する画像URL: " + imageUrl);
          console.log("📝 画像URLの長さ: " + imageUrl.length);
          
          sheet.getRange(eventRow, 9).setValue(imageUrl);
          
          // 更新後の値を確認
          var updatedImageUrl = sheet.getRange(eventRow, 9).getValue();
          console.log("📝 更新後の9列目の値: " + (updatedImageUrl || '(空)'));
          console.log("📝 更新後の9列目の値の長さ: " + (updatedImageUrl ? updatedImageUrl.length : 0));
          console.log("✅ スプレッドシートの画像URLを更新しました");
          
          // 念のため、もう一度確認
          var verifyImageUrl = sheet.getRange(eventRow, 9).getValue();
          console.log("📝 検証: 9列目の値: " + (verifyImageUrl || '(空)'));
          if (verifyImageUrl !== imageUrl) {
            console.error("❌ 警告: 9列目の値が期待値と異なります");
            console.error("❌ 期待値: " + imageUrl);
            console.error("❌ 実際の値: " + verifyImageUrl);
          } else {
            console.log("✅ 検証: 9列目の値が正しく設定されました");
          }
        }
        
        return output.setContent(JSON.stringify({ 
          status: 'success', 
          message: '画像をアップロードしました',
          imageUrl: imageUrl
        }));
      } catch (error) {
        console.error("❌ イベント画像アップロードエラー:", error);
        return output.setContent(JSON.stringify({ 
          status: 'error', 
          message: error.toString() 
        }));
      }
    }
    
    // ■パターンG：画像アップロード（JSON形式のBase64データ）
    else if (params.action === 'uploadImage' && params.image) {
      console.log("画像アップロード処理を開始（Base64形式）");
      
      try {
        var imageData = params.image;
        var fileName = params.fileName || 'event_' + Date.now() + '.jpg';
        
        // Base64デコード
        var decoded = Utilities.base64Decode(imageData);
        var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);
        
        // Google Driveに保存
        var folder = DriveApp.getFolderById(EVENT_IMAGE_FOLDER_ID);
        var file = folder.createFile(blob);
        
          // ファイルを公開設定にする
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          
          // Google Driveの画像を直接表示するためのURL形式に変換
          var fileId = file.getId();
          var imageUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
          
          console.log("画像アップロード成功: " + imageUrl);
          console.log("📁 ファイルID: " + fileId);
        
        return output.setContent(JSON.stringify({ 
          status: 'success', 
          imageUrl: imageUrl 
        }));
      } catch (error) {
        console.error("画像アップロードエラー:", error);
        return output.setContent(JSON.stringify({ 
          status: 'error', 
          message: error.toString() 
        }));
      }
    }

    // 未対応のformTypeの場合
    else {
      var errorMsg = "未対応のフォーム種別です: " + (params.formType || 'undefined') + ", action: " + (params.action || 'undefined');
      console.error("❌ エラー: " + errorMsg);
      console.error("params の全キー:", Object.keys(params));
      console.error("params の内容:", JSON.stringify(params));
      console.error("dataSource:", dataSource);
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
      sendEmailWithFrom(
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
// イベント取得関数
// ===========================================================

function getEvents(isArchived) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(EVENT_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return [];
    }
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return [];
    }
    
    // 現在の日時を取得
    var now = new Date();
    var nowStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    
    // ヘッダー行を読み取って列のインデックスを決定
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndexes = {
      id: 0,
      title: 1,
      category: 2,
      description: 3,
      date: 4,
      time: 5,
      capacity: 6,
      currentParticipants: 7,
      imageUrl: 8,
      archived: 9,
      createdAt: 10,
      updatedAt: 11,
      displayDeadline: 12,
      venue: 13,
      fee: 14,
      items: 15,
      registrant: 16
    };
    
    // ヘッダー行から列のインデックスを動的に決定
    for (var col = 0; col < headerRow.length; col++) {
      var headerValue = String(headerRow[col] || '').trim();
      if (headerValue === 'ID') columnIndexes.id = col;
      else if (headerValue === 'タイトル') columnIndexes.title = col;
      else if (headerValue === 'カテゴリ') columnIndexes.category = col;
      else if (headerValue === '説明') columnIndexes.description = col;
      else if (headerValue === '開催日') columnIndexes.date = col;
      else if (headerValue === '開催時間') columnIndexes.time = col;
      else if (headerValue === '定員') columnIndexes.capacity = col;
      else if (headerValue === '現在の参加者数') columnIndexes.currentParticipants = col;
      else if (headerValue === '画像URL') columnIndexes.imageUrl = col;
      else if (headerValue === 'アーカイブ') columnIndexes.archived = col;
      else if (headerValue === '作成日時') columnIndexes.createdAt = col;
      else if (headerValue === '更新日時') columnIndexes.updatedAt = col;
      else if (headerValue === '掲載期限') columnIndexes.displayDeadline = col;
      else if (headerValue === '会場') columnIndexes.venue = col;
      else if (headerValue === '参加費') columnIndexes.fee = col;
      else if (headerValue === '持ち物') columnIndexes.items = col;
      else if (headerValue === '登録者') columnIndexes.registrant = col;
    }
    
    console.log("📊 列インデックス:", JSON.stringify(columnIndexes));
    
    // より多くの列を取得（将来の拡張に対応）
    var maxColumn = Math.max(columnIndexes.id, columnIndexes.title, columnIndexes.category, 
                             columnIndexes.description, columnIndexes.date, columnIndexes.time,
                             columnIndexes.capacity, columnIndexes.currentParticipants, 
                             columnIndexes.imageUrl, columnIndexes.archived, columnIndexes.createdAt,
                             columnIndexes.updatedAt, columnIndexes.displayDeadline, 
                             columnIndexes.venue, columnIndexes.fee, columnIndexes.items, 
                             columnIndexes.registrant) + 1;
    var data = sheet.getRange(2, 1, lastRow - 1, maxColumn).getValues();
    var events = [];
    var seenIds = {}; // 重複チェック用
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      
      // 空の行をスキップ
      if (!row[0] || row[0] === '') {
        continue;
      }
      
      var eventId = row[columnIndexes.id];
      
      // 重複チェック
      if (seenIds[eventId]) {
        console.log("⚠️ 重複イベントをスキップ: ID=" + eventId);
        continue;
      }
      seenIds[eventId] = true;
      
      var archived = row[columnIndexes.archived] === true || row[columnIndexes.archived] === 'true' || row[columnIndexes.archived] === 'TRUE';
      
      // イベントの日時が過ぎていたら自動的にアーカイブ
      if (!archived && !isArchived) {
        var eventDate = row[columnIndexes.date];
        var eventTime = row[columnIndexes.time];
        
        if (eventDate && eventTime) {
          try {
            // 日付と時間を結合してDateオブジェクトを作成
            var dateStr = eventDate.toString();
            var timeStr = eventTime.toString();
            var timeParts = timeStr.split('-');
            
            if (timeParts.length >= 1) {
              var startTime = timeParts[0].trim();
              var eventDateTimeStr = dateStr + ' ' + startTime;
              var eventDateTime = new Date(eventDateTimeStr);
              
              // イベントの開始日時が現在より過去なら自動的にアーカイブ
              if (!isNaN(eventDateTime.getTime()) && eventDateTime < now) {
                console.log("⏰ イベントの日時が過ぎたため自動アーカイブ: ID=" + eventId + ", 日時=" + eventDateTimeStr);
                sheet.getRange(i, columnIndexes.archived + 1).setValue('true');
                archived = true;
              }
            }
          } catch (dateError) {
            console.error("日時の解析エラー: ID=" + eventId + ", エラー=" + dateError.toString());
          }
        }
      }
      
      if (archived === isArchived) {
        // Base64データをそのまま取得（Google Drive不要）
        var imageBase64 = row[columnIndexes.imageUrl] || null;
        var imageUrl = null;
        
        console.log("🖼️ スプレッドシートから取得した画像データ: " + (imageBase64 ? imageBase64.substring(0, 50) + '... (長さ: ' + imageBase64.length + ')' : '(空)'));
        console.log("🖼️ 画像データの型: " + typeof imageBase64);
        
        if (imageBase64 && typeof imageBase64 === 'string') {
          // Base64データかどうかを判定（長さが100文字以上で、httpで始まらない）
          if (imageBase64.length > 100 && !imageBase64.startsWith('http')) {
            // Base64データとしてそのまま使用（フロントエンドでdata URIに変換）
            imageUrl = imageBase64;
            console.log("🖼️ Base64データとして認識: " + imageBase64.substring(0, 50) + '...');
          }
          // 後方互換性: 古い形式のURLの場合はそのまま使用
          else if (imageBase64.indexOf('/file/d/') !== -1 || imageBase64.indexOf('uc?export=view&id=') !== -1 || imageBase64.indexOf('http') === 0) {
            imageUrl = imageBase64;
            console.log("🖼️ 古い形式のURLとして認識（後方互換性）: " + imageBase64);
          }
          // その他の場合はBase64データとして扱う
          else if (imageBase64.length > 50) {
            imageUrl = imageBase64;
            console.log("🖼️ Base64データとして扱う: " + imageBase64.substring(0, 50) + '...');
          }
        }
        
        console.log("🖼️ 最終的な画像データ: " + (imageUrl ? imageUrl.substring(0, 50) + '... (長さ: ' + imageUrl.length + ')' : '(空)'));
        
        // currentParticipantsの値を正しく取得（空文字列やnullの場合は0）
        var currentParticipants = 0;
        if (row[columnIndexes.currentParticipants] !== null && row[columnIndexes.currentParticipants] !== undefined && row[columnIndexes.currentParticipants] !== '') {
          currentParticipants = parseInt(row[columnIndexes.currentParticipants]) || 0;
        }
        
        // 会場の取得（登録者名が混入しないようにチェック）
        // デバッグ: 列の値を確認
        console.log("🔍 イベントID " + eventId + " の列データ確認:");
        console.log("  会場列インデックス: " + columnIndexes.venue + ", 値: " + (row[columnIndexes.venue] || '(空)'));
        console.log("  参加費列インデックス: " + columnIndexes.fee + ", 値: " + (row[columnIndexes.fee] || '(空)'));
        console.log("  持ち物列インデックス: " + columnIndexes.items + ", 値: " + (row[columnIndexes.items] || '(空)'));
        console.log("  登録者列インデックス: " + columnIndexes.registrant + ", 値: " + (row[columnIndexes.registrant] || '(空)'));
        
        var venueValue = row[columnIndexes.venue] || '';
        var registrantValue = row[columnIndexes.registrant] || '';
        var defaultVenue = 'pocopoco 東京都府中新町1丁目71-7';
        
        // 登録者名のリスト（会場に混入しないようにチェック用）
        var registrantNames = ['畠昂哉', '酒井くるみ', '平井菜央', '水石晶子', '長尾麻由子', '畠真姫'];
        
        // 会場が登録者名と同じ、または登録者名のリストに含まれている場合はデフォルト値を使用
        if (venueValue && venueValue.trim() !== '') {
          var trimmedVenue = venueValue.trim();
          if (trimmedVenue === registrantValue || registrantNames.indexOf(trimmedVenue) !== -1) {
            console.log("⚠️ 会場が登録者名と一致するため、デフォルト値を使用: " + trimmedVenue);
            venueValue = defaultVenue;
          }
        } else {
          venueValue = defaultVenue;
        }
        
        console.log("✅ 最終的な会場: " + venueValue);
        
        events.push({
          id: eventId,
          title: row[columnIndexes.title] || '',
          category: row[columnIndexes.category] || '',
          description: row[columnIndexes.description] || '',
          date: row[columnIndexes.date] || '',
          time: row[columnIndexes.time] || '',
          capacity: parseInt(row[columnIndexes.capacity]) || 0,
          currentParticipants: currentParticipants,
          imageUrl: imageUrl,
          imageBase64: imageBase64, // Base64データも含める
          displayDeadline: row[columnIndexes.displayDeadline] || '', // 掲載期限
          venue: venueValue, // 会場（登録者名が混入しないようにチェック済み）
          fee: row[columnIndexes.fee] || '無料', // 参加費
          items: row[columnIndexes.items] || '', // 持ち物
          // registrantはスタッフ間でのみ使用するため、フロントエンドには返さない
        });
      }
    }
    
    // 開催日でソート（古い順：開催予定は古い順、アーカイブは新しい順）
    if (isArchived) {
      events.sort(function(a, b) {
        return new Date(b.date) - new Date(a.date);
      });
    } else {
      events.sort(function(a, b) {
        return new Date(a.date) - new Date(b.date);
      });
    }
    
    return events;
  } catch (error) {
    console.error('イベント取得エラー:', error);
    return [];
  }
}

function getEventById(eventId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(EVENT_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() < 2) {
      return null;
    }
    
    // ヘッダー行を読み取って列のインデックスを決定
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndexes = {
      id: 0,
      title: 1,
      category: 2,
      description: 3,
      date: 4,
      time: 5,
      capacity: 6,
      currentParticipants: 7,
      imageUrl: 8,
      archived: 9,
      createdAt: 10,
      updatedAt: 11,
      displayDeadline: 12,
      venue: 13,
      fee: 14,
      items: 15,
      registrant: 16
    };
    
    // ヘッダー行から列のインデックスを動的に決定
    for (var col = 0; col < headerRow.length; col++) {
      var headerValue = String(headerRow[col] || '').trim();
      if (headerValue === 'ID') columnIndexes.id = col;
      else if (headerValue === 'タイトル') columnIndexes.title = col;
      else if (headerValue === 'カテゴリ') columnIndexes.category = col;
      else if (headerValue === '説明') columnIndexes.description = col;
      else if (headerValue === '開催日') columnIndexes.date = col;
      else if (headerValue === '開催時間') columnIndexes.time = col;
      else if (headerValue === '定員') columnIndexes.capacity = col;
      else if (headerValue === '現在の参加者数') columnIndexes.currentParticipants = col;
      else if (headerValue === '画像URL') columnIndexes.imageUrl = col;
      else if (headerValue === 'アーカイブ') columnIndexes.archived = col;
      else if (headerValue === '作成日時') columnIndexes.createdAt = col;
      else if (headerValue === '更新日時') columnIndexes.updatedAt = col;
      else if (headerValue === '掲載期限') columnIndexes.displayDeadline = col;
      else if (headerValue === '会場') columnIndexes.venue = col;
      else if (headerValue === '参加費') columnIndexes.fee = col;
      else if (headerValue === '持ち物') columnIndexes.items = col;
      else if (headerValue === '登録者') columnIndexes.registrant = col;
    }
    
    console.log("📊 getEventById 列インデックス:", JSON.stringify(columnIndexes));
    
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, columnIndexes.id + 1).getValue() === eventId) {
        // currentParticipantsの値を正しく取得（空文字列やnullの場合は0）
        var currentParticipants = 0;
        var participantsValue = sheet.getRange(i, columnIndexes.currentParticipants + 1).getValue();
        if (participantsValue !== null && participantsValue !== undefined && participantsValue !== '') {
          currentParticipants = parseInt(participantsValue) || 0;
        }
        
        var imageBase64 = sheet.getRange(i, columnIndexes.imageUrl + 1).getValue() || null;
        
        // 日付のフォーマット処理（Dateオブジェクトまたは文字列をYYYY-MM-DD形式に変換）
        var dateValue = sheet.getRange(i, columnIndexes.date + 1).getValue();
        var formattedDate = '';
        if (dateValue) {
          if (dateValue instanceof Date) {
            // Dateオブジェクトの場合
            var year = dateValue.getFullYear();
            var month = String(dateValue.getMonth() + 1).padStart(2, '0');
            var day = String(dateValue.getDate()).padStart(2, '0');
            formattedDate = year + '-' + month + '-' + day;
          } else {
            // 文字列の場合
            formattedDate = String(dateValue);
            // YYYY/MM/DD形式をYYYY-MM-DD形式に変換
            if (formattedDate.includes('/')) {
              formattedDate = formattedDate.replace(/\//g, '-');
            }
          }
        }
        
        // 掲載期限のフォーマット処理
        var deadlineValue = sheet.getRange(i, columnIndexes.displayDeadline + 1).getValue();
        var formattedDeadline = '';
        if (deadlineValue) {
          if (deadlineValue instanceof Date) {
            var year = deadlineValue.getFullYear();
            var month = String(deadlineValue.getMonth() + 1).padStart(2, '0');
            var day = String(deadlineValue.getDate()).padStart(2, '0');
            formattedDeadline = year + '-' + month + '-' + day;
          } else {
            formattedDeadline = String(deadlineValue);
            if (formattedDeadline.includes('/')) {
              formattedDeadline = formattedDeadline.replace(/\//g, '-');
            }
          }
        }
        
        return {
          id: sheet.getRange(i, columnIndexes.id + 1).getValue(),
          title: sheet.getRange(i, columnIndexes.title + 1).getValue(),
          category: sheet.getRange(i, columnIndexes.category + 1).getValue(),
          description: sheet.getRange(i, columnIndexes.description + 1).getValue(),
          date: formattedDate,
          time: sheet.getRange(i, columnIndexes.time + 1).getValue(),
          capacity: parseInt(sheet.getRange(i, columnIndexes.capacity + 1).getValue()) || 0,
          currentParticipants: currentParticipants,
          imageUrl: imageBase64,
          imageBase64: imageBase64,
          displayDeadline: formattedDeadline,
          venue: (function() {
            var venueValue = sheet.getRange(i, columnIndexes.venue + 1).getValue() || '';
            var registrantValue = sheet.getRange(i, columnIndexes.registrant + 1).getValue() || '';
            var defaultVenue = 'pocopoco 東京都府中新町1丁目71-7';
            
            // デバッグ: 列の値を確認
            console.log("🔍 イベントID " + eventId + " (getEventById) の列データ確認:");
            console.log("  会場列インデックス: " + (columnIndexes.venue + 1) + ", 値: " + (venueValue || '(空)'));
            console.log("  参加費列インデックス: " + (columnIndexes.fee + 1) + ", 値: " + (sheet.getRange(i, columnIndexes.fee + 1).getValue() || '(空)'));
            console.log("  持ち物列インデックス: " + (columnIndexes.items + 1) + ", 値: " + (sheet.getRange(i, columnIndexes.items + 1).getValue() || '(空)'));
            console.log("  登録者列インデックス: " + (columnIndexes.registrant + 1) + ", 値: " + (registrantValue || '(空)'));
            
            // 登録者名のリスト（会場に混入しないようにチェック用）
            var registrantNames = ['畠昂哉', '酒井くるみ', '平井菜央', '水石晶子', '長尾麻由子', '畠真姫'];
            
            // 会場が登録者名と同じ、または登録者名のリストに含まれている場合はデフォルト値を使用
            if (venueValue && venueValue.trim() !== '') {
              var trimmedVenue = venueValue.trim();
              if (trimmedVenue === registrantValue || registrantNames.indexOf(trimmedVenue) !== -1) {
                console.log("⚠️ 会場が登録者名と一致するため、デフォルト値を使用: " + trimmedVenue);
                return defaultVenue;
              }
            } else {
              return defaultVenue;
            }
            console.log("✅ 最終的な会場: " + venueValue);
            return venueValue;
          })(),
          fee: sheet.getRange(i, columnIndexes.fee + 1).getValue() || '無料',
          items: sheet.getRange(i, columnIndexes.items + 1).getValue() || '',
          registrant: sheet.getRange(i, columnIndexes.registrant + 1).getValue() || '' // 編集時のみ使用
        };
      }
    }
    
    return null;
  } catch (error) {
    console.error('イベント取得エラー:', error);
    return null;
  }
}

// ===========================================================
// 共通メール送信関数
// ===========================================================

// メール送信関数（送信元アドレスを指定可能）
function sendEmailWithFrom(to, subject, body, options) {
  // 受信者の検証
  if (!to || to.trim() === '') {
    console.error("✗ メール送信失敗: 受信者アドレスが指定されていません");
    return false;
  }
  
  try {
    // まずGmailAppを試行（エイリアス指定可能）
    try {
      var emailOptions = {
        name: FROM_NAME
      };
      
      // 送信元アドレスを指定（エイリアスが設定されている場合）
      if (FROM_EMAIL) {
        emailOptions.from = FROM_EMAIL;
      }
      
      // 追加オプションがあればマージ
      if (options) {
        for (var key in options) {
          emailOptions[key] = options[key];
        }
      }
      
      GmailApp.sendEmail(to, subject, body, emailOptions);
      console.log("✓ メール送信成功: " + subject + " (送信元: " + (emailOptions.from || "デフォルト") + ")");
      return true;
    } catch (gmailError) {
      // GmailAppが権限エラーの場合、MailAppにフォールバック
      console.log("⚠️ GmailApp送信に失敗（権限エラーの可能性）。MailAppで再試行します...");
      console.log("エラー詳細: " + gmailError.message);
      
      // MailAppで送信（エイリアスは実行アカウントの設定に依存）
      MailApp.sendEmail(to, subject, body);
      console.log("✓ メール送信成功（MailApp使用）: " + subject);
      console.log("ℹ️ 送信元は実行アカウントのメールアドレスになります");
      return true;
    }
  } catch (e) {
    console.error("✗ メール送信失敗: " + e.message);
    console.error("受信者: " + to);
    return false;
  }
}

function sendNotification(subject, body) {
  sendEmailWithFrom(NOTIFY_EMAIL, subject, body);
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
    sendEmailWithFrom(NOTIFY_EMAIL, "【最終テスト】ドライブアクセス成功", "GASはGoogleドライブへの書き込み権限を持っています。");
    console.log("✓ Drive権限テスト成功");
  } catch (e) {
    // 失敗した場合、権限がないかフォルダIDが間違っている
    sendEmailWithFrom(NOTIFY_EMAIL, "【最終テスト】⚠️ドライブアクセス失敗", "GASがドライブへのアクセス権を持っていません。エラー: " + e.message);
    console.error("✗ Drive権限テスト失敗:", e.message);
  }
}

// ===========================================================
// Gmail API権限を有効化するためのテスト関数
// ===========================================================
// この関数を実行すると、Gmail APIの権限要求ダイアログが表示されます
// 手順：
// 1. GASエディタでこの関数を選択
// 2. 「実行」ボタンをクリック
// 3. 権限要求ダイアログが表示されたら「許可」をクリック
// 4. これでGmailApp.sendEmailが使用可能になります

function enableGmailAPI() {
  try {
    console.log("🔄 Gmail APIの権限を要求します...");
    
    // GmailAppを使用して権限を要求（実際にはメールは送信しません）
    // このコードを実行すると、権限要求ダイアログが表示されます
    var testEmail = NOTIFY_EMAIL; // テスト用のメールアドレス
    
    // GmailApp.sendEmailを試行（権限がなければエラーになり、権限要求ダイアログが表示される）
    try {
      GmailApp.sendEmail(
        testEmail,
        "【Gmail API権限テスト】",
        "このメールはGmail APIの権限テストです。\n\n権限が正常に設定されていれば、このメールが送信されます。",
        {
          from: FROM_EMAIL,
          name: FROM_NAME
        }
      );
      console.log("✅ Gmail API権限が既に設定されています！");
      console.log("✅ メール送信テストが成功しました");
      return "成功: Gmail API権限が既に設定されています";
    } catch (e) {
      // 権限エラーの場合、権限要求ダイアログが表示される
      console.error("❌ Gmail API権限エラー:", e.message);
      console.log("ℹ️ 権限要求ダイアログが表示されたら、「許可」をクリックしてください");
      throw e; // エラーを再スローして、権限要求ダイアログを表示
    }
  } catch (e) {
    console.error("❌ エラー:", e.message);
    console.log("ℹ️ このエラーは正常です。権限要求ダイアログが表示されたら、「許可」をクリックしてください");
    throw e; // エラーを再スロー
  }
}
