// ChatGPT APIを使用して、M&A速報を分析する

// 設定

// API key
const INITIAL_API_KEY = "";

// model
// gpt-4o-mini: $0.15/1M tokens
const CHAT_MODEL = "gpt-4o-mini";


// 初期設定
const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

// ----------
// APIキーを設定する
// ソースにAPIキーを含めず、ScriptPropertiesに保存するためのメソッド
function setApiKey() {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
        PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', INITIAL_API_KEY);
    }
}
// APIキーの確認
function checkApiKey() {
    const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!apiKey) {
        throw new Error('APIキーが設定されていません。');
    }
    Logger.log("APIキーが設定されています。");
    Logger.log(apiKey);
}


// F列が空の行を探して、その行に対してclassifyItNewsByLineNumberを実行する
function classifyItNewsByLineNumberRange() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    for (let i = 1; i <= lastRow; i++) {
        // F列が空かどうかを確認(F列にはFalseという値が入っていることがあることに注意)
        const isItNewsValue = sheet.getRange(i, 6).getValue();
        // 空であるときに、classifyItNewsByLineNumberを実行する
        if (isItNewsValue === "") {  
            classifyItNewsByLineNumber(i, TARGET_SHEET_NAME);
        }
    }
}


// Spreadsheetの行番号の範囲指定で、IT業界関連かどうかを分類する
function isItNewsByLineNumberRange(startLineNumber, endLineNumber, sheetName) {
    // sheetNameのstartLineNumber行目からendLineNumber行目のタイトルを取得
    for (let i = startLineNumber; i <= endLineNumber; i++) {
        classifyItNewsByLineNumber(i, sheetName);
    }
}


// Spreadsheetの行番号指定で、IT業界関連かどうかを分類し、F列に結果を出力する
function classifyItNewsByLineNumber(lineNumber, sheetName) {
    // sheetNameのlineNumber行目のタイトルを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const title = sheet.getRange(lineNumber, 2).getValue();     // タイトルはB列(2列目)

    // タイトルが、IT業界関連かどうかを分類する
    const isItNewsValue = isItNews(title);

    // F列に結果を出力
    sheet.getRange(lineNumber, 6).setValue(isItNewsValue);     // F列(6列目)
}


// タイトルが、IT業界関連かどうかを分類する
// returns: [true: IT業界関連, false: IT業界関連でない]
function isItNews(newsTitle) {
    const payload = {
        model: CHAT_MODEL,
        messages: [
            {
                role: "system",
                content: "与えられたニュースタイトルがIT業界（SaaS, AI, クラウド, ソフトウェア開発, ITインフラ）関連か、そうでないかを分類してください。IT業界なら 'IT'、そうでないなら '非IT' と出力してください。"
            },
            {
                role: "user",
                content: newsTitle,
            },
        ],
    };
    const options = {
        method: "POST",
        contentType: "application/json",
        headers: {
            "Authorization": `Bearer ${OPENAI_API_KEY}`,
        },
        payload: JSON.stringify(payload),
    };

    // リクエストを送信
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);

    // レスポンスを解析
    const result = JSON.parse(response.getContentText());
    return result.choices[0].message.content === "IT";
}
