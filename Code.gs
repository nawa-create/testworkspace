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
 * スプレッドシートからデータを取得
 * ステータスが「チェック済み」以外のデータのみ取得
 */
function getData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('シート "' + SHEET_NAME + '" が見つかりません');
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return []; // ヘッダー行のみまたはデータなし
    }

    // 全データを取得（A列からK列まで）
    const range = sheet.getRange(2, 1, lastRow - 1, 11);
    const values = range.getValues();

    // データを整形してフィルタリング
    const data = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const status = row[10] ? row[10].toString().trim() : ''; // K列（ステータス）

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

    // 受信日時の降順でソート
    data.sort((a, b) => {
      const dateA = new Date(a.receivedDate);
      const dateB = new Date(b.receivedDate);
      return dateB - dateA; // 新しい順
    });

    return data;

  } catch (error) {
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
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('シート "' + SHEET_NAME + '" が見つかりません');
    }

    // K列（11列目）のステータスを更新
    sheet.getRange(rowIndex, 11).setValue(newStatus);

    return {
      success: true,
      message: 'ステータスを更新しました'
    };

  } catch (error) {
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
