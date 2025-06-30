/**
 * @fileoverview 注文DB(D_Order)のアーカイブ処理を担当します。
 * @author T.Maruyama
 * @since 2025-06-30
 * @version 1.2.2
 */

/**
 * =================================================================================
 * 変更履歴
 * =================================================================================
 * 2025-06-30 T.Maruyama v1.2.2
 * - 新規作成
 * =================================================================================
 */

/**
 * このスクリプトを使用する前に、Webアプリ側のプロジェクトをライブラリとして追加してください。
 * ライブラリの識別子は、以下のコード内で 'WebAppLib' としています。
 * 実際の識別子に合わせて修正してください。
 */

function archiveOrders() {
  // WebAppLib は、ライブラリを追加する際に設定した識別子
  const lib = WebAppLib; 

  try {
    // --- 処理開始前にアクセスを制限 ---
    lib.setMaintenanceMode(true);
    Logger.log('Webアプリのメンテナンスモードを開始しました。');

    // ▼設定項目
    const sourceSheetName = 'D_Order';
    // ope
    // const archiveFolderId = '1Otc8zc2-ycXfywZji56dCb5-8GGVfU7_';
    // const backupFolderId = '1Otc8zc2-ycXfywZji56dCb5-8GGVfU7_';    
    // dev
    const archiveFolderId = '1q5vh7bGb2oXdC6UZyJtAS51HN1K6h1yM';
    const backupFolderId = '1q5vh7bGb2oXdC6UZyJtAS51HN1K6h1yM';
    const notificationEmail = 't-maruyama@ito-ex.co.jp';
    const slackWebhookUrl = 'https://hooks.slack.com/services/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';
    // ▲設定項目

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const spreadsheetName = ss.getName();
    let summaryMessage = '';

    // --- 1. スプレッドシートのバックアップを作成 ---
    const originalFile = DriveApp.getFileById(ss.getId());
    const backupFolder = DriveApp.getFolderById(backupFolderId);
    const backupFileName = `[バックアップ]${spreadsheetName}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    originalFile.makeCopy(backupFileName, backupFolder);
    Logger.log(`スプレッドシートのバックアップを作成しました: ${backupFileName}`);

    // --- 2. アーカイブ対象データの抽出 ---
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      throw new Error(`シート「${sourceSheetName}」が見つかりません。`);
    }

    const today = new Date();
    // 2ヶ月前の16日をアーカイブの基準日とする
    const archiveBorderDate = new Date(today.getFullYear(), today.getMonth() - 2, 16);
    archiveBorderDate.setHours(0, 0, 0, 0);

    const dataRange = sourceSheet.getDataRange();
    const allData = dataRange.getValues();
    const header = allData.shift(); // ヘッダー行を取得

    const rowsToArchive = [];
    const rowsToDelete = [];
    const orderDateColumnIndex = 1;  // B列 (注文日)
    const updateDateColumnIndex = 7; // H列 (更新日時)

    allData.forEach((row, index) => {
      // 注文日が存在する行のみを対象
      if (row[orderDateColumnIndex]) {
        const orderDate = new Date(row[orderDateColumnIndex]);
        if (orderDate < archiveBorderDate) {
          // CSV用に日付フォーマットを修正
          const newRow = [...row];
          newRow[orderDateColumnIndex] = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
          if (row[updateDateColumnIndex]) {
            newRow[updateDateColumnIndex] = Utilities.formatDate(new Date(row[updateDateColumnIndex]), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
          }
          rowsToArchive.push(newRow);
          // 削除対象の行番号を記録 (2行目から始まるため index + 2)
          rowsToDelete.push(index + 2);
        }
      }
    });

    // --- 3. CSVとZIPの作成、シートからのデータ削除 ---
    if (rowsToArchive.length === 0) {
      summaryMessage = `[${spreadsheetName}] アーカイブ対象のデータはありませんでした。\n\n処理は正常に終了しました。`;
      Logger.log(summaryMessage);
    } else {
      const csvContent = [header, ...rowsToArchive].map(r => r.join(',')).join('\n');
      const csvFileName = `D_Order_Archive_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')}.csv`;
      const csvBlob = Utilities.newBlob(csvContent, MimeType.CSV, csvFileName);
      const zipFileName = `D_Order_Archive_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}.zip`;
      const zipBlob = Utilities.zip([csvBlob], zipFileName);
      
      const archiveFolder = DriveApp.getFolderById(archiveFolderId);
      archiveFolder.createFile(zipBlob);
      Logger.log(`${zipFileName} をGoogle Driveに保存しました。`);

      // 行を削除（下の行から削除しないとインデックスがずれる）
      rowsToDelete.reverse().forEach(rowNum => {
        sourceSheet.deleteRow(rowNum);
      });
      Logger.log(`${rowsToDelete.length} 件のデータを元のシートから削除しました。`);

      summaryMessage = `[${spreadsheetName}] バックアップを作成後、${rowsToArchive.length}件のデータをZIPファイル（${zipFileName}）としてアーカイブし、元のシートから削除しました。\n\n処理は正常に終了しました。`;
    }

    // --- 4. 処理完了通知 ---
    sendNotifications(`✅ [GAS 完了通知] ${spreadsheetName}`, summaryMessage, notificationEmail, slackWebhookUrl);

  } catch (e) {
    Logger.log(e);
    const errorMessage = `[${spreadsheetName}] 処理中にエラーが発生しました。\n\nエラー内容:\n${e.message}\n\nスクリプトのログを確認してください。`;
    // --- エラー通知 ---
    sendNotifications(`❌ [GAS エラー通知] ${spreadsheetName}`, errorMessage, notificationEmail, slackWebhookUrl);
  } finally {
    // --- 処理終了後に必ずアクセスを再開 ---
    lib.setMaintenanceMode(false);
    Logger.log('Webアプリのメンテナンスモードを終了しました。');
  }
}


/**
 * メールとSlackに通知を送信するヘルパー関数
 * @param {string} subject - メールの件名
 * @param {string} body - メッセージ本文
 * @param {string} email - 通知先メールアドレス
 * @param {string} slackUrl - SlackのWebhook URL
 */
function sendNotifications(subject, body, email, slackUrl) {
  // メール送信
  if (email) {
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
    });
  }
  
  // Slack送信
  if (slackUrl && slackUrl.startsWith('https://hooks.slack.com/')) {
    const payload = {
      "text": `*${subject}*\n\n${body}`
    };
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    };
    UrlFetchApp.fetch(slackUrl, options);
  }
}
