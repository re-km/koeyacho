/**
 * 橋梁点検 音声入力システム - GASバックエンド
 *
 * 【機能】
 * 1. 指定フォルダ内のスプレッドシートを「橋リスト」として取得（サブフォルダ含む）
 * 2. 選択された橋（ファイルID）にデータを書き込み
 * 3. シート（部位+径間）がなければ自動作成
 * 4. Gemini で音声誤変換を補正し、参照表から損傷番号を解決
 *
 * 【Script Properties】
 * - GEMINI_API_KEY: Gemini API キー
 * - LOOKUP_SPREADSHEET_ID: 参照用 Google スプレッドシート ID
 */

// ▼▼▼ 設定エリア ▼▼▼
const TARGET_FOLDER_ID = "1EUOTznh8cuaz_5o-EPC2plCJVxxgNnG4";
const LOOKUP_SHEET_NAME = "部材リスト、損傷リスト";
const LOOKUP_CACHE_SECONDS = 6 * 60 * 60;
const PART_KEYS = ["トラス・アーチ", "下面", "下部工", "支承部", "橋面"];
const PART_MEMBER_INDEXES = {
    "下面": 0,
    "下部工": 1,
    "支承部": 2,
    "橋面": 3,
    "トラス・アーチ": 4
};
const DAMAGE_ID_INDEX = 9;
const DAMAGE_NAME_INDEX = 10;
const CRACK_DAMAGE_NAMES = ["ひびわれ", "床版ひびわれ", "亀裂"];
const MEMBER_CORRECTIONS = [
    { pattern: /評判/g, replacement: "床版" },
    { pattern: /商番/g, replacement: "床版" },
    { pattern: /常磐/g, replacement: "床版" },
    { pattern: /床板/g, replacement: "床版" },
    { pattern: /主げた/g, replacement: "主桁" },
    { pattern: /主ケタ/g, replacement: "主桁" },
    { pattern: /横げた/g, replacement: "横桁" },
    { pattern: /縦げた/g, replacement: "縦桁" }
];
const DAMAGE_CORRECTIONS = [
    { pattern: /床版ひび割れ/g, replacement: "床版ひびわれ" },
    { pattern: /ひび割れ/g, replacement: "ひびわれ" },
    { pattern: /クラック/g, replacement: "ひびわれ" },
    { pattern: /剥離・?鉄筋露出/g, replacement: "剥離・鉄筋露出" },
    { pattern: /鉄筋露出/g, replacement: "剥離・鉄筋露出" },
    { pattern: /漏水・?遊離石灰/g, replacement: "漏水・遊離石灰" },
    { pattern: /遊離石灰/g, replacement: "漏水・遊離石灰" }
];
const FALLBACK_MEMBER_NAMES = ["床版", "主桁", "横桁", "縦桁", "支承本体", "高欄"];
// ▲▲▲ 設定エリア ▲▲▲

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const fileId = data.fileId;
        if (!fileId) {
            throw new Error("ファイルIDが指定されていません");
        }

        const ss = SpreadsheetApp.openById(fileId);
        const sheetName = data.sheetName || "点検データ";
        let sheet = ss.getSheetByName(sheetName);

        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            setupSheetHeader(sheet);
        }

        let parsedData;
        let warnings = [];
        let normalizedText = "";
        let resolvedDamageId = "";

        if (data.rawText) {
            const parseResult = parseInspectionRawText(data.rawText, sheetName);
            parsedData = parseResult.record;
            warnings = parseResult.warnings;
            normalizedText = parseResult.normalizedText;
            resolvedDamageId = parseResult.resolvedDamageId;
        } else {
            const legacyResult = buildRecordFromLegacyPayload(data, sheetName);
            parsedData = legacyResult.record;
            warnings = legacyResult.warnings;
            resolvedDamageId = parsedData.damageId || "";
        }

        const nextNo = sheet.getLastRow();
        const newRow = [
            nextNo,
            parsedData.prevRecord || "",
            parsedData.sameAngle || "",
            parsedData.photoNo || "",
            parsedData.emergency || "",
            parsedData.member || "",
            parsedData.material || "",
            parsedData.elementNumber || "",
            parsedData.damageId || "",
            parsedData.degree || "",
            parsedData.crackSpacing || "",
            parsedData.crackWidth || "",
            parsedData.dimensions || "",
            parsedData.judgment || "",
            parsedData.progress || "",
            parsedData.thirdParty || "",
            parsedData.remarks || ""
        ];

        writeInspectionRow(sheet, nextNo + 1, newRow);

        return createJsonResponse({
            success: true,
            message: `${ss.getName()} > ${sheetName} に保存しました`,
            sheetName: sheetName,
            assignedNo: nextNo,
            parseStatus: warnings.length ? "warning" : "ok",
            warnings: warnings,
            normalizedText: normalizedText || "",
            resolvedDamageId: resolvedDamageId || ""
        });
    } catch (error) {
        return createJsonResponse({ success: false, error: error.toString() });
    }
}

function parseInspectionRawText(rawText, sheetName) {
    const warnings = [];
    const partKey = getPartKey(sheetName);
    if (!partKey) {
        warnings.push(`部位を判定できませんでした: ${sheetName}`);
    }

    let lookupData = null;
    try {
        lookupData = getLookupData();
    } catch (lookupError) {
        console.error("Lookup sheet error:", lookupError);
        warnings.push("参照表を読み込めなかったため番号解決を一部スキップしました");
    }

    const memberRows = lookupData && partKey ? lookupData.memberRowsByPart[partKey] || [] : [];
    const damageRows = lookupData ? lookupData.damageRows : [];
    let geminiResult = {
        normalizedText: rawText,
        memberCandidate: "",
        damageCandidate: "",
        unresolvedTerms: []
    };

    try {
        geminiResult = callGeminiNormalizer(rawText, partKey, memberRows, damageRows);
    } catch (geminiError) {
        console.error("Gemini API Error:", geminiError);
        warnings.push(`Gemini補正に失敗したためローカル解析で処理しました: ${summarizeGeminiError(geminiError)}`);
    }

    const normalizedText = normalizeRecognizedText(geminiResult.normalizedText || rawText);
    const record = buildParsedRecord(normalizedText, partKey, memberRows, damageRows, geminiResult, warnings);

    return {
        record: record,
        warnings: dedupeStrings(warnings),
        normalizedText: normalizedText,
        resolvedDamageId: record.damageId || ""
    };
}

function buildRecordFromLegacyPayload(data, sheetName) {
    const warnings = [];
    const record = createEmptyRecord();
    const partKey = getPartKey(sheetName);

    record.prevRecord = safeTrim(data.prevRecord);
    record.sameAngle = safeTrim(data.sameAngle);
    record.photoNo = safeTrim(data.photoNo);
    record.emergency = safeTrim(data.emergency);
    record.material = safeTrim(data.material);
    record.elementNumber = safeTrim(data.elementNumber);
    record.degree = safeTrim(data.degree);
    record.crackSpacing = safeTrim(data.crackSpacing);
    record.crackWidth = safeTrim(data.crackWidth);
    record.dimensions = safeTrim(data.dimensions);
    record.judgment = safeTrim(data.judgment);
    record.progress = safeTrim(data.progress);
    record.thirdParty = safeTrim(data.thirdParty);

    let lookupData = null;
    try {
        lookupData = getLookupData();
    } catch (lookupError) {
        console.error("Lookup sheet error:", lookupError);
        warnings.push("参照表を読み込めなかったため番号解決を一部スキップしました");
    }

    const memberRows = lookupData && partKey ? lookupData.memberRowsByPart[partKey] || [] : [];
    const damageRows = lookupData ? lookupData.damageRows : [];

    const resolvedMember = resolveMemberValue(safeTrim(data.member), safeTrim(data.member), memberRows);
    if (resolvedMember) {
        record.member = resolvedMember;
    } else if (safeTrim(data.member)) {
        record.member = memberRows.length ? "" : safeTrim(data.member);
        warnings.push(`部材を正式名称へ正規化できませんでした: ${safeTrim(data.member)}`);
    }

    const rawDamageId = safeTrim(data.damageId);
    if (/^\d+$/.test(rawDamageId)) {
        record.damageId = rawDamageId;
    } else if (rawDamageId) {
        const resolvedDamage = resolveDamageValue(rawDamageId, rawDamageId, damageRows);
        if (resolvedDamage) {
            record.damageId = resolvedDamage.id;
        } else {
            warnings.push(`損傷番号に変換できませんでした: ${rawDamageId}`);
        }
    }

    record.remarks = joinRemarks([
        safeTrim(data.remarks),
        warnings.length ? `要確認: ${dedupeStrings(warnings).join(" / ")}` : ""
    ]);

    return {
        record: record,
        warnings: dedupeStrings(warnings)
    };
}

function buildParsedRecord(normalizedText, partKey, memberRows, damageRows, geminiResult, warnings) {
    const record = createEmptyRecord();
    const workingText = normalizedText || "";

    record.prevRecord = extractPrevRecord(workingText);
    record.photoNo = extractPhotoNo(workingText);
    record.degree = extractDegree(workingText);

    const resolvedMember = resolveMemberValue(
        geminiResult.memberCandidate,
        workingText,
        memberRows
    );
    if (resolvedMember) {
        record.member = resolvedMember;
    } else if (memberRows.length) {
        const hintedMember = safeTrim(geminiResult.memberCandidate) || extractMemberHint(workingText, memberRows);
        if (hintedMember) {
            warnings.push(`部材を正式名称へ正規化できませんでした: ${hintedMember}`);
        }
    }

    const resolvedDamage = resolveDamageValue(
        geminiResult.damageCandidate,
        workingText,
        damageRows
    );
    let resolvedDamageName = "";
    if (resolvedDamage) {
        record.damageId = resolvedDamage.id;
        resolvedDamageName = resolvedDamage.name;
    } else {
        const hintedDamage = extractDamageHint(workingText, damageRows) || safeTrim(geminiResult.damageCandidate);
        if (hintedDamage) {
            resolvedDamageName = hintedDamage;
            warnings.push(`損傷番号に変換できませんでした: ${hintedDamage}`);
        }
    }

    record.crackWidth = extractCrackWidth(workingText, resolvedDamageName);
    record.dimensions = extractDimensions(workingText, resolvedDamageName);
    record.elementNumber = extractElementNumber(workingText, record.photoNo, record.dimensions);

    const unresolvedTerms = normalizeStringArray(geminiResult.unresolvedTerms);
    if (unresolvedTerms.length) {
        warnings.push(`未解析語: ${unresolvedTerms.join("、")}`);
    }

    record.remarks = warnings.length ? `要確認: ${dedupeStrings(warnings).join(" / ")}` : "";
    return record;
}

function createEmptyRecord() {
    return {
        prevRecord: "",
        sameAngle: "",
        photoNo: "",
        emergency: "",
        member: "",
        material: "",
        elementNumber: "",
        damageId: "",
        degree: "",
        crackSpacing: "",
        crackWidth: "",
        dimensions: "",
        judgment: "",
        progress: "",
        thirdParty: "",
        remarks: ""
    };
}

function getLookupData() {
    const lookupSpreadsheetId = PropertiesService.getScriptProperties().getProperty("LOOKUP_SPREADSHEET_ID");
    if (!lookupSpreadsheetId) {
        throw new Error("LOOKUP_SPREADSHEET_ID is not set in Script Properties.");
    }

    const cacheKey = `lookup-data:v2:${lookupSpreadsheetId}`;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
        return JSON.parse(cached);
    }

    const ss = SpreadsheetApp.openById(lookupSpreadsheetId);
    const sheet = ss.getSheetByName(LOOKUP_SHEET_NAME);
    if (!sheet) {
        throw new Error(`${LOOKUP_SHEET_NAME} シートが見つかりません`);
    }

    const memberRowsByPart = {};
    Object.keys(PART_MEMBER_INDEXES).forEach((partKey) => {
        memberRowsByPart[partKey] = [];
    });

    const damageRows = [];
    const memberSeen = {};
    const damageSeen = {};
    Object.keys(PART_MEMBER_INDEXES).forEach((partKey) => {
        memberSeen[partKey] = {};
    });

    const lastRow = sheet.getLastRow();
    if (lastRow >= 3) {
        const values = sheet.getRange(3, 2, lastRow - 2, 11).getDisplayValues();
        values.forEach((row) => {
            Object.keys(PART_MEMBER_INDEXES).forEach((partKey) => {
                const memberName = safeTrim(row[PART_MEMBER_INDEXES[partKey]]);
                if (memberName && !memberSeen[partKey][memberName]) {
                    memberRowsByPart[partKey].push(createLookupItem(memberName, "", "member"));
                    memberSeen[partKey][memberName] = true;
                }
            });

            const damageId = safeTrim(row[DAMAGE_ID_INDEX]);
            const damageName = safeTrim(row[DAMAGE_NAME_INDEX]);
            if (damageId && damageName && !damageSeen[damageName]) {
                damageRows.push(createLookupItem(damageName, damageId, "damage"));
                damageSeen[damageName] = true;
            }
        });
    }

    Object.keys(memberRowsByPart).forEach((partKey) => {
        memberRowsByPart[partKey].sort(sortLookupItemsBySpecificity);
    });
    damageRows.sort(sortLookupItemsBySpecificity);

    const lookupData = {
        memberRowsByPart: memberRowsByPart,
        damageRows: damageRows
    };

    cache.put(cacheKey, JSON.stringify(lookupData), LOOKUP_CACHE_SECONDS);
    return lookupData;
}

function createLookupItem(name, id, type) {
    return {
        id: id || "",
        name: name,
        normalizedName: type === "damage" ? normalizeDamageMatchText(name) : normalizeMemberMatchText(name)
    };
}

function sortLookupItemsBySpecificity(left, right) {
    return right.normalizedName.length - left.normalizedName.length;
}

function callGeminiNormalizer(rawText, partKey, memberRows, damageRows) {
    const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        throw new Error("GEMINI_API_KEY is not set in Script Properties.");
    }

    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    const prompt = buildGeminiPrompt(rawText, partKey, memberRows, damageRows);
    const payload = {
        contents: [{
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: {
            responseMimeType: "application/json",
            temperature: 0.1
        }
    };
    const options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const resultJson = JSON.parse(responseText);
    if (response.getResponseCode() !== 200) {
        throw new Error(`Gemini API request failed: ${responseText}`);
    }

    const content = (((resultJson.candidates || [])[0] || {}).content || {}).parts || [];
    const text = content.map((part) => part.text || "").join("").trim();
    if (!text) {
        throw new Error("Gemini response was empty");
    }

    const parsed = parseLooseJson(text);
    return {
        normalizedText: safeTrim(parsed.normalizedText) || rawText,
        memberCandidate: safeTrim(parsed.memberCandidate),
        damageCandidate: safeTrim(parsed.damageCandidate),
        unresolvedTerms: normalizeStringArray(parsed.unresolvedTerms)
    };
}

function buildGeminiPrompt(rawText, partKey, memberRows, damageRows) {
    const memberCandidates = memberRows.map((row) => row.name);
    const damageCandidates = damageRows.map((row) => row.name);
    return [
        "あなたは橋梁点検の音声補正アシスタントです。",
        "iPhone 音声認識の誤変換を、橋梁点検で使う正式名称に補正してください。",
        `対象部位: ${partKey || "不明"}`,
        `候補部材: ${memberCandidates.length ? memberCandidates.join("、") : "候補なし"}`,
        `候補損傷: ${damageCandidates.length ? damageCandidates.join("、") : "候補なし"}`,
        "",
        "ルール:",
        "1. normalizedText は元の語順をなるべく維持し、誤変換補正と表記正規化だけを行う。",
        "2. 部材名は候補部材から、損傷名は候補損傷から選ぶ。部分的な発話は正式名称へ寄せる。",
        "3. 例: 評判→床版、鉄筋露出→剥離・鉄筋露出、ひび割れ→ひびわれ。",
        "4. 写真番号や要素番号などの数字は推測で変更しない。",
        "5. 乗算記号は * に統一する。",
        "6. unresolvedTerms には補正できなかった語だけを配列で入れる。",
        "",
        "出力は次の JSON のみ。",
        '{"normalizedText":"","memberCandidate":"","damageCandidate":"","unresolvedTerms":[]}',
        "",
        `入力音声: ${rawText}`
    ].join("\n");
}

function parseLooseJson(text) {
    const cleaned = text
        .replace(/^\s*```(?:json)?/i, "")
        .replace(/```\s*$/i, "")
        .trim();
    return JSON.parse(cleaned);
}

function resolveMemberValue(candidateText, normalizedText, memberRows) {
    const hints = dedupeStrings([
        safeTrim(candidateText),
        extractMemberHint(normalizedText, memberRows),
        extractFallbackMemberHint(normalizedText)
    ]);

    for (let i = 0; i < hints.length; i += 1) {
        const match = findLookupMatch(hints[i], memberRows, normalizeMemberMatchText);
        if (match) {
            return match.name;
        }
    }

    if (!memberRows.length) {
        return hints.length ? hints[0] : "";
    }
    return "";
}

function resolveDamageValue(candidateText, normalizedText, damageRows) {
    const hints = dedupeStrings([
        extractDamageHint(normalizedText, damageRows),
        safeTrim(candidateText)
    ]);

    let bestMatch = null;
    for (let i = 0; i < hints.length; i += 1) {
        const match = findLookupMatch(hints[i], damageRows, normalizeDamageMatchText);
        if (match && (!bestMatch || match.normalizedName.length > bestMatch.normalizedName.length)) {
            bestMatch = match;
        }
    }
    return bestMatch;
}

function findLookupMatch(text, rows, normalizer) {
    const normalizedText = normalizer(text);
    if (!normalizedText || !rows.length) {
        return null;
    }

    for (let i = 0; i < rows.length; i += 1) {
        if (rows[i].normalizedName === normalizedText) {
            return rows[i];
        }
    }

    for (let i = 0; i < rows.length; i += 1) {
        if (rows[i].normalizedName.indexOf(normalizedText) !== -1 || normalizedText.indexOf(rows[i].normalizedName) !== -1) {
            return rows[i];
        }
    }

    return null;
}

function extractMemberHint(text, memberRows) {
    if (!text) {
        return "";
    }

    const normalizedText = normalizeMemberMatchText(text);
    for (let i = 0; i < memberRows.length; i += 1) {
        if (normalizedText.indexOf(memberRows[i].normalizedName) !== -1) {
            return memberRows[i].name;
        }
    }
    return "";
}

function extractFallbackMemberHint(text) {
    const corrected = applyMemberCorrections(text || "");
    const normalizedText = normalizeMemberMatchText(corrected);
    for (let i = 0; i < FALLBACK_MEMBER_NAMES.length; i += 1) {
        const memberName = FALLBACK_MEMBER_NAMES[i];
        if (normalizedText.indexOf(normalizeMemberMatchText(memberName)) !== -1) {
            return memberName;
        }
    }
    return "";
}

function extractDamageHint(text, damageRows) {
    if (!text) {
        return "";
    }

    const normalizedText = normalizeDamageMatchText(text);
    for (let i = 0; i < damageRows.length; i += 1) {
        if (normalizedText.indexOf(damageRows[i].normalizedName) !== -1) {
            return damageRows[i].name;
        }
    }

    const corrected = applyDamageCorrections(text);
    if (corrected !== text) {
        const correctedNormalizedText = normalizeDamageMatchText(corrected);
        for (let i = 0; i < damageRows.length; i += 1) {
            if (correctedNormalizedText.indexOf(damageRows[i].normalizedName) !== -1) {
                return damageRows[i].name;
            }
        }
    }

    return "";
}

function normalizeRecognizedText(text) {
    let value = safeTrim(text).normalize("NFKC");
    value = value.replace(/([0-9.]+)\s*[×xX＊]\s*([0-9.]+)/g, "$1*$2");
    value = value.replace(/\u3000/g, " ");
    value = value.replace(/[‐‑‒–—―−]/g, "-");
    value = applyMemberCorrections(value);
    value = applyDamageCorrections(value);
    value = value.replace(/\s+/g, " ").trim();
    return value;
}

function applyMemberCorrections(text) {
    return applyCorrections(text, MEMBER_CORRECTIONS);
}

function applyDamageCorrections(text) {
    return applyCorrections(text, DAMAGE_CORRECTIONS);
}

function applyCorrections(text, replacements) {
    let value = safeTrim(text);
    replacements.forEach((rule) => {
        value = value.replace(rule.pattern, rule.replacement);
    });
    return value;
}

function normalizeMemberMatchText(text) {
    return normalizeMatchText(applyMemberCorrections(text));
}

function normalizeDamageMatchText(text) {
    return normalizeMatchText(applyDamageCorrections(text));
}

function normalizeMatchText(text) {
    return safeTrim(text)
        .normalize("NFKC")
        .toLowerCase()
        .replace(/([0-9.]+)\s*[×xX＊]\s*([0-9.]+)/g, "$1*$2")
        .replace(/[‐‑‒–—―−]/g, "-")
        .replace(/[・･、，,。]/g, "")
        .replace(/[()（）［］\[\]{}｛｝]/g, "")
        .replace(/\s+/g, "");
}

function extractPrevRecord(text) {
    const match = safeTrim(text).match(/前回(?:調書)?\s*([0-9A-Za-z\-]+)/);
    return match ? match[1] : "";
}

function extractPhotoNo(text) {
    const patterns = [
        /写真(?:番号)?\s*([0-9０-９]{1,4})/i,
        /([0-9０-９]{1,4})\s*番?\s*写真/i
    ];
    for (let i = 0; i < patterns.length; i += 1) {
        const match = safeTrim(text).match(patterns[i]);
        if (match) {
            return normalizeDigits(match[1]);
        }
    }
    return "";
}

function extractDegree(text) {
    const match = safeTrim(text).match(/(?:程度)\s*([a-eA-E]|[1-5]-[c-eC-E])/i);
    return match ? match[1].toLowerCase() : "";
}

function extractCrackWidth(text, damageName) {
    const workingText = safeTrim(text);
    const explicitPatterns = [
        /(?:ひび(?:われ|割れ)幅|クラック幅|ひび幅|幅)\s*([0-9０-９]+(?:\.[0-9０-９]+)?)(?![0-9.]*\*)(?:\s*(mm|㎜|ミリ))?/i,
        /(?:ひび(?:われ|割れ)|クラック|亀裂)[^0-9]{0,6}([0-9０-９]+(?:\.[0-9０-９]+)?)(?![0-9.]*\*)(?:\s*(mm|㎜|ミリ))?/i
    ];

    for (let i = 0; i < explicitPatterns.length; i += 1) {
        const match = workingText.match(explicitPatterns[i]);
        if (match) {
            return `${toCleanNumberString(toNumber(match[1]))}mm`;
        }
    }

    if (!isWidthBasedCrackDamage(damageName)) {
        return "";
    }

    const pair = extractDimensionPair(workingText);
    if (pair) {
        return `${toCleanNumberString(toNumber(pair.first))}mm`;
    }

    const implicit = workingText.match(/([0-9０-９]+(?:\.[0-9０-９]+)?)(?![0-9.]*\*)\s*(mm|㎜|ミリ)/i);
    if (implicit) {
        return `${toCleanNumberString(toNumber(implicit[1]))}mm`;
    }

    return "";
}

function extractDimensions(text, damageName) {
    const pair = extractDimensionPair(text);
    if (pair) {
        const unit = normalizeDimensionUnit(pair.unitText, pair.first, pair.second);
        if (isWidthBasedCrackDamage(damageName)) {
            return formatDimensionValue(pair.second, unit);
        }
        return `${formatDimensionValue(pair.first, unit)}*${formatDimensionValue(pair.second, unit)}`;
    }

    const singleMatch = safeTrim(text).match(/(?:数量|長さ|延長|面積|範囲|広さ)\s*([0-9０-９]+(?:\.[0-9０-９]+)?)(?:\s*(mm|㎜|ミリ|m|ｍ|メートル))?/i);
    if (singleMatch) {
        const unit = normalizeDimensionUnit(singleMatch[2], singleMatch[1]);
        return formatDimensionValue(singleMatch[1], unit);
    }

    return "";
}

function extractDimensionPair(text) {
    const pairMatch = safeTrim(text).match(/([0-9０-９]+(?:\.[0-9０-９]+)?)\s*\*\s*([0-9０-９]+(?:\.[0-9０-９]+)?)(?:\s*(mm|㎜|ミリ|m|ｍ|メートル))?/i);
    if (!pairMatch) {
        return null;
    }

    return {
        first: pairMatch[1],
        second: pairMatch[2],
        unitText: pairMatch[3] || ""
    };
}

function isDeckCrackDamage(damageName) {
    const normalizedDamage = normalizeDamageMatchText(damageName || "");
    return normalizedDamage.indexOf(normalizeDamageMatchText("床版ひびわれ")) !== -1;
}

function isWidthBasedCrackDamage(damageName) {
    const normalizedDamage = normalizeDamageMatchText(damageName || "");
    return CRACK_DAMAGE_NAMES.some((name) => normalizedDamage.indexOf(normalizeDamageMatchText(name)) !== -1);
}

function normalizeDimensionUnit(unitText, firstValue, secondValue) {
    const normalizedUnit = safeTrim(unitText).toLowerCase();
    if (normalizedUnit === "mm" || normalizedUnit === "㎜" || normalizedUnit === "ミリ") {
        return "mm";
    }
    if (normalizedUnit === "m" || normalizedUnit === "ｍ" || normalizedUnit === "メートル") {
        return "m";
    }
    return inferDimensionUnit(firstValue, secondValue);
}

function inferDimensionUnit(firstValue, secondValue) {
    const first = safeTrim(firstValue);
    const second = safeTrim(secondValue);
    const hasDecimal = first.indexOf(".") !== -1 || second.indexOf(".") !== -1;
    if (hasDecimal) {
        return "m";
    }
    const firstNumber = toNumber(first);
    const secondNumber = second ? toNumber(second) : 0;
    return firstNumber >= 10 || secondNumber >= 10 ? "mm" : "m";
}

function formatDimensionValue(valueText, unit) {
    const numericValue = toNumber(valueText);
    if (!isFinite(numericValue)) {
        return "";
    }
    return unit === "mm"
        ? toCleanNumberString(numericValue / 1000)
        : toCleanNumberString(numericValue);
}

function extractElementNumber(text, photoNo, dimensions) {
    let working = safeTrim(text);
    if (photoNo) {
        working = working
            .replace(new RegExp(`写真(?:番号)?\\s*${escapeRegExp(photoNo)}`, "gi"), " ")
            .replace(new RegExp(`${escapeRegExp(photoNo)}\\s*番?\\s*写真`, "gi"), " ");
    }
    if (dimensions) {
        const escapedDimensions = escapeRegExp(dimensions);
        working = working.replace(new RegExp(escapedDimensions, "g"), " ");
    }
    working = working.replace(/([0-9０-９]+(?:\.[0-9０-９]+)?)\s*\*\s*([0-9０-９]+(?:\.[0-9０-９]+)?)(?:\s*(mm|㎜|ミリ|m|ｍ|メートル))?/gi, " ");

    const explicitMatch = working.match(/(?:要素(?:番号)?|部材番号|番号)\s*([0-9０-９]{1,4})/);
    if (explicitMatch) {
        return normalizeDigits(explicitMatch[1]);
    }

    const genericMatch = working.match(/(?:^|[^0-9])([0-9０-９]{4})(?=[^0-9]|$)/);
    if (genericMatch) {
        return normalizeDigits(genericMatch[1]);
    }

    return "";
}

function normalizeDigits(value) {
    return safeTrim(value).normalize("NFKC").replace(/[^\d]/g, "");
}

function toNumber(value) {
    return Number(safeTrim(value).normalize("NFKC").replace(/,/g, ""));
}

function toCleanNumberString(value) {
    const rounded = Number(Number(value).toFixed(6));
    return isFinite(rounded) ? String(rounded) : "";
}

function getPartKey(sheetName) {
    const currentSheetName = safeTrim(sheetName);
    for (let i = 0; i < PART_KEYS.length; i += 1) {
        if (currentSheetName.indexOf(PART_KEYS[i]) === 0) {
            return PART_KEYS[i];
        }
    }
    return "";
}

function normalizeStringArray(value) {
    if (!value) {
        return [];
    }
    if (Array.isArray(value)) {
        return dedupeStrings(value.map((item) => safeTrim(item)).filter(Boolean));
    }
    if (typeof value === "string") {
        return dedupeStrings(value.split(/[、,\n]/).map((item) => safeTrim(item)).filter(Boolean));
    }
    return [];
}

function joinRemarks(segments) {
    return dedupeStrings((segments || []).map((segment) => safeTrim(segment)).filter(Boolean)).join(" / ");
}

function dedupeStrings(values) {
    const result = [];
    const seen = {};
    (values || []).forEach((value) => {
        const key = safeTrim(value);
        if (key && !seen[key]) {
            seen[key] = true;
            result.push(key);
        }
    });
    return result;
}

function safeTrim(value) {
    return value == null ? "" : String(value).trim();
}

function summarizeGeminiError(error) {
    const message = safeTrim(error && error.message ? error.message : error);
    if (!message) {
        return "理由不明";
    }

    return message
        .replace(/\s+/g, " ")
        .slice(0, 180);
}

function escapeRegExp(text) {
    return safeTrim(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function doGet(e) {
    try {
        if (TARGET_FOLDER_ID === "ここにフォルダIDを貼り付け") {
            return createJsonResponse({
                status: "error",
                message: "GAS側でフォルダIDが設定されていません。スクリプトを確認してください。"
            });
        }

        const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
        const bridgeList = [];
        getAllFilesRecursively(folder, bridgeList);
        bridgeList.sort((a, b) => a.name.localeCompare(b.name, "ja"));

        return createJsonResponse({
            status: "success",
            bridges: bridgeList,
            folderName: folder.getName()
        });
    } catch (error) {
        return createJsonResponse({ status: "error", error: error.toString() });
    }
}

function getAllFilesRecursively(folder, list) {
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

    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
        getAllFilesRecursively(subFolders.next(), list);
    }
}

function writeInspectionRow(sheet, rowNumber, rowValues) {
    ensureTextColumns(sheet);
    sheet.getRange(rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
}

function ensureTextColumns(sheet) {
    sheet.getRangeList(["B:B", "C:C", "D:D", "H:H"]).setNumberFormat("@");
}

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

    sheet.setColumnWidth(1, 50);
    sheet.setColumnWidth(2, 60);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(9, 120);
    sheet.setColumnWidth(13, 80);
    sheet.setColumnWidth(17, 200);
    ensureTextColumns(sheet);
}

function createJsonResponse(data) {
    return ContentService
        .createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}






