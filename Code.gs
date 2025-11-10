// スプレッドシートのID
const SPREADSHEET_ID = '1sT_UQFSmWBuUza4H1Xp1nmgWYqS0WbvcuhjhSjtNvE8';
const SHEET_NAME = 'FAX受信データ';

/**
 * Webアプリのエントリーポイント
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FAX受信データ検証システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 接続テスト用の関数
 * クライアントから呼び出せるかをテスト
 */
function testConnection() {
  Logger.log('=== testConnection() 呼び出し成功 ===');
  return {
    success: true,
    message: 'サーバーとの接続に成功しました！',
    timestamp: new Date().toLocaleString('ja-JP')
  };
}

/**
 * スプレッドシートからデータを取得
 * ステータスが「チェック済み」以外のデータのみ取得
 */
function getData() {
  try {
    Logger.log('=== getData() 開始 ===');
    Logger.log('スプレッドシートID: ' + SPREADSHEET_ID);
    Logger.log('シート名: ' + SHEET_NAME);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('スプレッドシート取得成功');

    const sheet = ss.getSheetByName(SHEET_NAME);
    Logger.log('シート取得結果: ' + (sheet ? '成功' : '失敗'));

    if (!sheet) {
      Logger.log('エラー: シートが見つかりません');
      throw new Error('シート "' + SHEET_NAME + '" が見つかりません');
    }

    const lastRow = sheet.getLastRow();
    Logger.log('最終行: ' + lastRow);

    if (lastRow <= 1) {
      Logger.log('データなし（ヘッダー行のみ）');
      return []; // ヘッダー行のみまたはデータなし
    }

    // 全データを取得（A列からK列まで）
    const range = sheet.getRange(2, 1, lastRow - 1, 11);
    Logger.log('取得範囲: A2:K' + lastRow);

    const values = range.getValues();
    Logger.log('取得した行数: ' + values.length);

    // データを整形してフィルタリング
    const data = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const status = row[10] ? row[10].toString().trim() : ''; // K列（ステータス）

      Logger.log('行' + (i + 2) + ': ステータス=' + status);

      // ステータスが「チェック済み」以外のデータのみ追加
      if (status !== 'チェック済み') {
        data.push({
          rowIndex: i + 2, // スプレッドシートの実際の行番号（ヘッダー分+1）
          receivedDate: row[0] ? formatDate(row[0]) : '',
          customerName: row[1] || '',
          deliveryDestination: row[2] || '',
          requester: row[3] || '',
          deliveryDate: row[4] || '',
          productName: row[5] || '',
          packaging: row[6] || '',
          quantity: row[7] || '',
          remarks: row[8] || '',
          imageUrl: row[9] || '',
          status: status || '未チェック'
        });
      }
    }

    Logger.log('フィルタ後のデータ件数: ' + data.length);

    // 受信日時の降順でソート
    data.sort((a, b) => {
      const dateA = new Date(a.receivedDate);
      const dateB = new Date(b.receivedDate);
      return dateB - dateA; // 新しい順
    });

    Logger.log('ソート完了');
    Logger.log('=== getData() 正常終了 ===');

    return data;

  } catch (error) {
    Logger.log('=== getData() エラー ===');
    Logger.log('エラーメッセージ: ' + error.message);
    Logger.log('エラースタック: ' + error.stack);
    console.error('データ取得エラー:', error);
    throw new Error('データの取得に失敗しました: ' + error.message);
  }
}

/**
 * 日付のフォーマット
 */
function formatDate(date) {
  if (!date) return '';

  // 既に文字列の場合はそのまま返す
  if (typeof date === 'string') return date;

  // Date型の場合はフォーマット
  if (date instanceof Date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    const seconds = String(date.getSeconds()).padStart(2, '0');
    return `${year}/${month}/${day} ${hours}:${minutes}:${seconds}`;
  }

  return date.toString();
}

/**
 * ステータスを更新
 * @param {number} rowIndex - スプレッドシートの行番号
 * @param {string} newStatus - 新しいステータス
 */
function updateStatus(rowIndex, newStatus) {
  try {
    Logger.log('=== updateStatus() 開始 ===');
    Logger.log('行番号: ' + rowIndex);
    Logger.log('新しいステータス: ' + newStatus);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      Logger.log('エラー: シートが見つかりません');
      throw new Error('シート "' + SHEET_NAME + '" が見つかりません');
    }

    // K列（11列目）のステータスを更新
    sheet.getRange(rowIndex, 11).setValue(newStatus);
    Logger.log('ステータス更新成功');
    Logger.log('=== updateStatus() 正常終了 ===');

    return {
      success: true,
      message: 'ステータスを更新しました'
    };

  } catch (error) {
    Logger.log('=== updateStatus() エラー ===');
    Logger.log('エラーメッセージ: ' + error.message);
    Logger.log('エラースタック: ' + error.stack);
    console.error('ステータス更新エラー:', error);
    throw new Error('ステータスの更新に失敗しました: ' + error.message);
  }
}

/**
 * Google DriveのファイルIDを抽出
 * @param {string} url - Google DriveのURL
 * @return {string} - ファイルID
 */
function extractDriveFileId(url) {
  if (!url) return '';

  // https://drive.google.com/file/d/FILE_ID/view?usp=drivesdk 形式から抽出
  const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : '';
}
