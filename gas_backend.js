/**
 * 橋梁点検 音声入力システム - GASバックエンド（フォルダ参照型・サブフォルダ対応版）
 * 
 * 【機能】
 * 1. 指定フォルダ内のスプレッドシートを自動で「橋リスト」として取得（サブフォルダ含む）
 * 2. 選択された橋（ファイルID）にデータを書き込み
 * 3. シート（部位+径間）がなければ自動作成
 * 
 * 【設定手順】
 * 1. Google Driveに「点検データ」などのフォルダを作成
 * 2. そのフォルダのURL末尾のID（folders/の後ろの文字列）をコピー
 * 3. 以下の TARGET_FOLDER_ID に貼り付け
 * 4. デプロイ → ウェブアプリ → 全員アクセスで公開
 */

// ▼▼▼ 設定エリア ▼▼▼
// ここに点検データを保存する親フォルダのIDを貼り付けてください
const TARGET_FOLDER_ID = "ここにフォルダIDを貼り付け";
// ▲▲▲ 設定エリア ▲▲▲

// POSTリクエスト：データ書き込み
function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const fileId = data.fileId; // 書き込み対象のファイルID

        // ファイルIDが指定されていない場合
        if (!fileId) {
            throw new Error("ファイルIDが指定されていません");
        }

        // スプレッドシートを取得
        const ss = SpreadsheetApp.openById(fileId);

        // シート名を決定（部位+径間）
        const sheetName = data.sheetName || "点検データ";
        let sheet = ss.getSheetByName(sheetName);

        // シートがなければ作成
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            setupSheetHeader(sheet);
        }

        // データを追加
        const timestamp = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
        const newRow = [
            timestamp,
            data.id || "",
            data.member || "",
            data.damageId || "",
            data.damageName || "",
            data.dimensions || "",
            data.timestamp || ""
        ];

        sheet.appendRow(newRow);

        return createJsonResponse({
            success: true,
            message: `${ss.getName()} > ${sheetName} に保存しました`,
            sheetName: sheetName
        });

    } catch (error) {
        return createJsonResponse({ success: false, error: error.toString() });
    }
}

// GETリクエスト：橋リスト（ファイル一覧）の取得
function doGet(e) {
    try {
        // フォルダIDが未設定の場合
        if (TARGET_FOLDER_ID === "ここにフォルダIDを貼り付け") {
            return createJsonResponse({
                status: "error",
                message: "GAS側でフォルダIDが設定されていません。スクリプトを確認してください。"
            });
        }

        const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
        const bridgeList = [];

        // サブフォルダも含めて全ファイルを探索
        getAllFilesRecursively(folder, bridgeList);

        // 名前順にソート
        bridgeList.sort((a, b) => a.name.localeCompare(b.name, 'ja'));

        return createJsonResponse({
            status: "success",
            bridges: bridgeList,
            folderName: folder.getName()
        });

    } catch (error) {
        return createJsonResponse({ status: "error", error: error.toString() });
    }
}

// 再帰的にファイルを探索する関数
function getAllFilesRecursively(folder, list) {
    // 1. 直下のスプレッドシートを取得
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
        const file = files.next();
        list.push({
            name: file.getName(),
            id: file.getId(),
            url: file.getUrl(),
            lastUpdated: Utilities.formatDate(file.getLastUpdated(), "Asia/Tokyo", "yyyy/MM/dd HH:mm")
        });
    }

    // 2. サブフォルダがあれば潜る
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        getAllFilesRecursively(subFolders.next(), list);
    }
}

// ヘッダー設定（共通関数）
function setupSheetHeader(sheet) {
    sheet.getRange(1, 1, 1, 7).setValues([
        ["受信日時", "No", "部材", "変状番号", "変状名", "寸法", "入力時刻"]
    ]);

    const header = sheet.getRange(1, 1, 1, 7);
    header.setBackground("#4285f4");
    header.setFontColor("white");
    header.setFontWeight("bold");
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 120);
}

// JSONレスポンス生成（CORS対応）
function createJsonResponse(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}
