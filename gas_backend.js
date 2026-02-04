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
        // データを追加
        // No. は現在の行数（ヘッダー分引く必要なし？ いいえ、ヘッダーが1行目なので、現在データ行数+1 ＝ getLastRow()でよい）
        // 例: ヘッダーのみ(1行) -> getLastRow=1 -> 次は No.1
        // 例: データ1件あり(2行) -> getLastRow=2 -> 次は No.2
        const nextNo = sheet.getLastRow();

        const newRow = [
            nextNo,                 // No.
            data.prevRecord || "",  // 前回調書
            "",                     // 同ｱﾝｸﾞﾙ写
            data.photoNo || "",     // 写真番号
            "",                     // 応急措置写真
            data.member || "",      // 部材
            "",                     // 材料
            data.elementNumber || "", // 要素番号
            data.damageId || "",    // 変状 (番号で入力)
            data.degree || "",      // 程度
            data.crackSpacing || "",// ひび間隔
            data.crackWidth || "",  // ひび幅
            data.dimensions || "",  // 数量(m)
            "",                     // 判定
            "",                     // 進行
            "",                     // 第三者被害
            data.remarks || ""      // 備考
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
    const headers = [
        "No.", "前回調書", "同ｱﾝｸﾞﾙ写", "写真番号", "応急措置写真",
        "部材", "材料", "要素番号", "変状", "程度",
        "ひび間隔", "ひび幅", "数量(m)", "判定", "進行",
        "第三者被害", "備考"
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("white");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);

    // 列幅調整 (主な項目のみ)
    sheet.setColumnWidth(1, 50);  // No.
    sheet.setColumnWidth(2, 60);  // 前回
    sheet.setColumnWidth(4, 60);  // 写真番号
    sheet.setColumnWidth(6, 120); // 部材
    sheet.setColumnWidth(9, 120); // 変状
    sheet.setColumnWidth(13, 80); // 数量
    sheet.setColumnWidth(17, 200); // 備考
}

// JSONレスポンス生成（CORS対応）
function createJsonResponse(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}
