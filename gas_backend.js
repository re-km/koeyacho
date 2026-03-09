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

        // URL末尾のIDやフォルダIDのフェッチなどは既存処理を活用
        const ss = SpreadsheetApp.openById(fileId);

        // シート名を決定（部位+径間）
        const sheetName = data.sheetName || "点検データ";
        let sheet = ss.getSheetByName(sheetName);

        // シートがなければ作成
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            setupSheetHeader(sheet);
        }

        // --- [NEW] Gemini API 解析処理 ---
        let parsedData = {};
        if (data.rawText) {
            try {
                parsedData = callGeminiAPI(data.rawText);
            } catch (geminiError) {
                console.error("Gemini API Error:", geminiError);
                // 失敗時は rawText を備考に突っ込むフォールバック
                parsedData = { remarks: "【APIエラー】" + data.rawText };
            }
        } else {
            // 万が一、古いバージョンからの送信等で rawText が無い場合のフォールバック
            parsedData = {
                member: data.member,
                damageId: data.damageId,
                crackWidth: data.crackWidth,
                dimensions: data.dimensions,
                photoNo: data.photoNo,
                prevRecord: data.prevRecord,
                remarks: data.remarks
            };
        }

        // データを追加
        const nextNo = sheet.getLastRow();

        const newRow = [
            nextNo,                        // No.
            parsedData.prevRecord || "",   // 前回調書
            parsedData.sameAngle || "",    // 同ｱﾝｸﾞﾙ写
            parsedData.photoNo || "",      // 写真番号
            parsedData.emergency || "",    // 応急措置写真
            parsedData.member || "",       // 部材
            parsedData.material || "",     // 材料
            parsedData.elementNumber || "",// 要素番号
            parsedData.damageId || "",     // 変状 (番号または名前)
            parsedData.degree || "",       // 程度
            parsedData.crackSpacing || "", // ひび間隔
            parsedData.crackWidth || "",   // ひび幅
            parsedData.dimensions || "",   // 数量(m)
            parsedData.judgment || "",     // 判定
            parsedData.progress || "",     // 進行
            parsedData.thirdParty || "",   // 第三者被害
            parsedData.remarks || ""       // 備考
        ];

        sheet.appendRow(newRow);

        return createJsonResponse({
            success: true,
            message: `${ss.getName()} > ${sheetName} に保存しました`,
            sheetName: sheetName,
            assignedNo: nextNo
        });

    } catch (error) {
        return createJsonResponse({ success: false, error: error.toString() });
    }
}

// --- [NEW] Gemini API 呼び出し関数 ---
function callGeminiAPI(rawText) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        throw new Error("GEMINI_API_KEY is not set in Script Properties.");
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    // プロンプト設計 (system_prompt.md 準拠 + スプレッドシート列対応)
    const prompt = `あなたは橋梁点検の専門アシスタントです。現場の点検員が話す音声を解析し、スプレッドシートの指定列に対応するJSONデータに変換してください。

【入力音声】
${rawText}

【抽出・変換ルール】
1. 以下のJSONスキーマに従って出力してください。該当しない項目は空文字("")にしてください。
2. 音声認識の誤変換（「商番」→「床版」、「夕刊」→「遊間」など）は文脈から推測して正しい橋梁用語・変状名に補正してください。
3. ひび幅は「幅0.2」などなら crackWidth="0.2mm" とし、メートル換算はしないでください。
4. 数量（長さ、面積など）はメートル単位に換算して dimensions に入れてください（例: 500ミリ→0.5）。
5. 抽出できなかった不明な単語や呟きは、すべて remarks(備考) に入れてください。

【出力フォーマット要求】
必ず以下のキーを持つ純粋なJSON文字列(バッククォート等のマークダウン不要)のみを出力してください。
{
  "prevRecord": "前回調書番号",
  "sameAngle": "同アングル写真",
  "photoNo": "写真番号",
  "emergency": "応急措置写真",
  "member": "部材名（主桁、床版、支承など正しい用語で）",
  "material": "材料",
  "elementNumber": "要素番号",
  "damageId": "変状名（ひび割れ、剥離・鉄筋露出など）または番号",
  "degree": "程度（a,b,cなど）",
  "crackSpacing": "ひび間隔",
  "crackWidth": "ひび幅（単位変換なし）",
  "dimensions": "数量（メートル換算した数値）",
  "judgment": "判定",
  "progress": "進行",
  "thirdParty": "第三者被害",
  "remarks": "備考やその他の発言内容"
}`;

    const payload = {
        "contents": [{
            "parts": [{
                "text": prompt
            }]
        }],
        "generationConfig": {
            "responseMimeType": "application/json",
            "temperature": 0.1
        }
    };

    const options = {
        "method": "post",
        "contentType": "application/json",
        "payload": JSON.stringify(payload),
        "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(url, options);
    const resultJson = JSON.parse(response.getContentText());

    if (response.getResponseCode() !== 200) {
        throw new Error("Gemini API request failed: " + JSON.stringify(resultJson));
    }

    try {
        const text = resultJson.candidates[0].content.parts[0].text;
        return JSON.parse(text);
    } catch (e) {
        throw new Error("Failed to parse Gemini response as JSON: " + e.message);
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
