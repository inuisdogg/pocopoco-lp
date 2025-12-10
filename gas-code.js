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
    
    return output.setContent(JSON.stringify({ status: 'error', message: '無効なアクション' }));
  } catch (error) {
    console.error('doGetエラー:', error);
    return output.setContent(JSON.stringify({ status: 'error', message: error.toString() }));
  }
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
    console.log("e.postData の存在:", !!(e && e.postData));
    console.log("e.parameter の存在:", !!(e && e.parameter));
    
    // パラメータ取得の試行
    var params = {};
    var dataSource = "";
    
    // 方法1: postData.contentsから取得（JSON形式）
    if (e && e.postData && e.postData.contents) {
      console.log("📥 postData.contents から取得を試行");
      console.log("postData.type:", e.postData.type);
      console.log("postData.contents の長さ:", e.postData.contents.length);
      console.log("postData.contents の最初の200文字:", e.postData.contents.substring(0, 200));
      
      try {
        params = JSON.parse(e.postData.contents);
        dataSource = "postData.contents";
        console.log("✅ postData.contents から取得成功");
        console.log("取得したparams.action:", params.action);
        console.log("取得したparams.formType:", params.formType);
        
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
          var imageUrl = file.getUrl().replace('/file/d/', '/uc?export=view&id=').replace('/view?usp=sharing', '');
          
          console.log("画像アップロード成功: " + imageUrl);
          
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
          'その他メッセージ'
        ]);
        // ヘッダー行を太字にする
        var headerRange = sheet.getRange(1, 1, 1, 12);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#E8E8E8');
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
        "■参加人数: " + (params.participants || '') + "\n\n";
      
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
      
      console.log("イベント申し込みフォーム: 処理完了");
    }

    // ■パターンE：イベント管理（作成・更新・削除・アーカイブ）
    else if (params.action === 'createEvent' || params.action === 'updateEvent' || params.action === 'deleteEvent' || params.action === 'archiveEvent') {
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
          'ID',
          'タイトル',
          'カテゴリ',
          '説明',
          '開催日',
          '開催時間',
          '定員',
          '現在の参加者数',
          '画像URL',
          'アーカイブ',
          '作成日時',
          '更新日時'
        ]);
        // ヘッダー行を太字にする
        var headerRange = sheet.getRange(1, 1, 1, 12);
        headerRange.setFontWeight('bold');
        headerRange.setBackground('#E8E8E8');
        console.log("✅ ヘッダー行を追加しました");
      } else {
        console.log("✅ 既存のシートを使用します");
        console.log("📊 シートの最終行: " + sheet.getLastRow());
      }
      
      var date = new Date();
      var dateStr = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      
      // 画像がある場合は先にアップロード
      var imageUrl = params.imageUrl || '';
      if (params.imageBase64) {
        console.log("画像アップロード処理を開始");
        console.log("画像データの長さ: " + (params.imageBase64 ? params.imageBase64.length : 0));
        try {
          var decoded = Utilities.base64Decode(params.imageBase64);
          var fileName = params.imageFileName || 'event_' + Date.now() + '.jpg';
          var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);
          console.log("Blob作成完了: " + fileName);
          
          var folder = DriveApp.getFolderById(EVENT_IMAGE_FOLDER_ID);
          console.log("フォルダ取得完了: " + EVENT_IMAGE_FOLDER_ID);
          var file = folder.createFile(blob);
          console.log("ファイル作成完了: " + file.getName());
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          imageUrl = file.getUrl().replace('/file/d/', '/uc?export=view&id=').replace('/view?usp=sharing', '');
          
          console.log("画像アップロード成功: " + imageUrl);
        } catch (error) {
          console.error("画像アップロードエラー:", error);
          console.error("エラースタック:", error.stack);
        }
      } else {
        console.log("画像データなし");
      }
      
      // イベント作成
      if (params.action === 'createEvent') {
        // 最大IDを取得
        var lastRow = sheet.getLastRow();
        var maxId = 0;
        if (lastRow > 1) {
          var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
          for (var i = 0; i < ids.length; i++) {
            if (ids[i][0] > maxId) {
              maxId = ids[i][0];
            }
          }
        }
        var newId = maxId + 1;
        
        var rowData = [
          newId,
          params.title || '',
          params.category || '',
          params.description || '',
          params.date || '',
          params.time || '',
          params.capacity || 0,
          params.currentParticipants || 0,
          imageUrl,
          'false',
          dateStr,
          dateStr
        ];
        
        console.log("📝 追加する行データ: " + JSON.stringify(rowData));
        console.log("📝 追加前の最終行: " + sheet.getLastRow());
        
        sheet.appendRow(rowData);
        
        console.log("✅ イベント作成完了: ID=" + newId);
        console.log("📝 追加後の最終行: " + sheet.getLastRow());
        console.log("📝 追加した行の内容を確認:");
        var addedRow = sheet.getRange(sheet.getLastRow(), 1, 1, 12).getValues()[0];
        console.log("📝 " + JSON.stringify(addedRow));
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
            // 新しい画像がある場合は更新、ない場合は既存の画像URLを保持
            if (imageUrl) {
              sheet.getRange(i, 9).setValue(imageUrl);
            }
            sheet.getRange(i, 12).setValue(dateStr);
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
          MailApp.sendEmail(
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

    // ■パターンF：画像アップロード（JSON形式のBase64データ）
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
        var imageUrl = file.getUrl().replace('/file/d/', '/uc?export=view&id=').replace('/view?usp=sharing', '');
        
        console.log("画像アップロード成功: " + imageUrl);
        
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
    var data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    var events = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var archived = row[9] === true || row[9] === 'true';
      
      if (archived === isArchived) {
        events.push({
          id: row[0],
          title: row[1],
          category: row[2],
          description: row[3],
          date: row[4],
          time: row[5],
          capacity: row[6],
          currentParticipants: row[7] || 0,
          imageUrl: row[8] || null
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
    
    var lastRow = sheet.getLastRow();
    
    for (var i = 2; i <= lastRow; i++) {
      if (sheet.getRange(i, 1).getValue() === eventId) {
        return {
          id: sheet.getRange(i, 1).getValue(),
          title: sheet.getRange(i, 2).getValue(),
          category: sheet.getRange(i, 3).getValue(),
          description: sheet.getRange(i, 4).getValue(),
          date: sheet.getRange(i, 5).getValue(),
          time: sheet.getRange(i, 6).getValue(),
          capacity: sheet.getRange(i, 7).getValue(),
          currentParticipants: sheet.getRange(i, 8).getValue() || 0,
          imageUrl: sheet.getRange(i, 9).getValue() || null
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
