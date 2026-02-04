/**
 * 橋梁点検用スプレッドシート一括作成ツール
 * 
 * 【使い方】
 * 1. このスクリプトが含まれるスプレッドシートに「一括作成リスト」という名前のシートを作成します。
 * 2. 1行目をヘッダーとして、2行目以降に以下のデータを入力してください：
 *    - A列: 橋の名前（必須） 例: ○○橋
 *    - B列: フォルダ名（任意） 例: R5年度A班
 * 
 * 3. 上部メニューの「一括作成」→「シート作成開始」を実行します。
 *    ※初回は権限承認が必要です。
 * 
 * 4. 指定したフォルダ（TARGET_FOLDER_ID）の中にファイルが作成されます。
 *    フォルダ名が指定されている場合は、サブフォルダが自動作成（または既存使用）され、その中に格納されます。
 */

// ▼▼▼ 設定エリア ▼▼▼
// ファイル作成先の親フォルダID（gas_backend.jsと同じIDを設定してください）
const TARGET_FOLDER_ID = "ここにフォルダIDを貼り付け";
// ▲▲▲ 設定エリア ▲▲▲

// メニューを追加
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🌉 一括作成')
        .addItem('シート作成開始', 'createSheetsFromList')
        .addToUi();
}

// 一括作成のメイン処理
function createSheetsFromList() {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const listSheet = ss.getSheetByName("一括作成リスト");

    // フォルダIDチェック
    if (TARGET_FOLDER_ID === "ここにフォルダIDを貼り付け") {
        ui.alert("エラー", "スクリプト内の TARGET_FOLDER_ID にフォルダIDを設定してください。", ui.ButtonSet.OK);
        return;
    }

    // シートチェック
    if (!listSheet) {
        const newSheet = ss.insertSheet("一括作成リスト");
        newSheet.getRange("A1:C1").setValues([["橋の名前(必須)", "フォルダ名(任意)", "作成結果"]]);
        newSheet.getRange("A1:C1").setBackground("#4285f4").setFontColor("white").setFontWeight("bold");
        newSheet.getRange("A2").setValue("例：○○橋");
        newSheet.setColumnWidth(1, 200);
        newSheet.setColumnWidth(2, 150);
        newSheet.setColumnWidth(3, 300);

        ui.alert("準備", "「一括作成リスト」シートを作成しました。\nA列に橋名を入力して、もう一度実行してください。", ui.ButtonSet.OK);
        return;
    }

    // データ取得（A列:橋名, B列:フォルダ名）
    const lastRow = listSheet.getLastRow();
    if (lastRow < 2) {
        ui.alert("データなし", "作成するデータがありません。", ui.ButtonSet.OK);
        return;
    }

    const values = listSheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const resultRange = listSheet.getRange(2, 3, lastRow - 1, 1);
    const results = [];

    // 親フォルダ取得
    const parentFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);

    // サブフォルダのキャッシュ（何度も取得しないように）
    const folderCache = {};

    let createCount = 0;
    let skipCount = 0;

    // ループ処理
    for (let i = 0; i < values.length; i++) {
        const bridgeName = values[i][0];
        const subFolderName = values[i][1];

        if (!bridgeName) {
            results.push(["スキップ: 名前なし"]);
            skipCount++;
            continue;
        }

        try {
            // 保存先フォルダの決定
            let targetFolder = parentFolder;

            if (subFolderName) {
                if (!folderCache[subFolderName]) {
                    // サブフォルダを探す、なければ作る
                    const folders = parentFolder.getFoldersByName(subFolderName);
                    if (folders.hasNext()) {
                        folderCache[subFolderName] = folders.next();
                    } else {
                        folderCache[subFolderName] = parentFolder.createFolder(subFolderName);
                    }
                }
                targetFolder = folderCache[subFolderName];
            }

            // 同じ名前のファイルがあるかチェック
            const existingFiles = targetFolder.getFilesByName(bridgeName);
            if (existingFiles.hasNext()) {
                results.push(["スキップ: 同名ファイルあり"]);
                skipCount++;
                continue;
            }

            // スプレッドシート作成
            const newSS = SpreadsheetApp.create(bridgeName); // ルートに作られる
            const file = DriveApp.getFileById(newSS.getId());
            file.moveTo(targetFolder); // 指定フォルダに移動

            // 基本シート設定
            setupSheet(newSS);

            results.push([`作成完了 (${targetFolder.getName()})`]);
            createCount++;

        } catch (e) {
            results.push([`エラー: ${e.toString()}`]);
        }
    }

    // 結果書き込み
    resultRange.setValues(results);

    ui.alert("完了", `処理が終了しました。\n作成: ${createCount}件\nスキップ: ${skipCount}件`, ui.ButtonSet.OK);
}

// 作成したシートの初期設定
function setupSheet(ss) {
    // デフォルトのシート1の名前を変更
    let sheet = ss.getSheets()[0];
    sheet.setName("下面"); // デフォルトで「下面」シートにしておく

    // ヘッダー追加
    const header = ["受信日時", "No", "部材", "変状番号", "変状名", "寸法", "入力時刻"];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);

    const headerRange = sheet.getRange(1, 1, 1, header.length);
    headerRange.setBackground("#4285f4");
    headerRange.setFontColor("white");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(5, 150);
    sheet.setColumnWidth(6, 120);
}
