// ページ全体から、解説（買い手の戦略、目的）300文字くらいでまとめる


function testSummarizePageByLineNumberRange() {
    summarizePageByLineNumberRange(TARGET_SHEET_NAME);
}

// 指定したシートの全行に対して、summarizePageByLineNumberを実行する
function summarizePageByLineNumberRange(sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
        summarizePageByLineNumber(i, sheetName);
    }
}


// 行番号を指定して、summarizePageを実行し、その結果をH列に設定する
// ただし、G列の値が入っていて、かつ、その値が".pdf"で終わっていない場合のみ、実行する
function summarizePageByLineNumber(lineNumber) {
    const title = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME).getRange(lineNumber, 2).getValue();
    Logger.log(`分析中: ${title}`);

    const url = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME).getRange(lineNumber, 7).getValue();
    if (!url || url.endsWith(".pdf")) {
        Logger.log(`  └ PDFのためスキップ`);
        return;
    }
    
    // 分析済みだったらスキップ
    if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME).getRange(lineNumber, 8).getValue()) {
        Logger.log(`  └ 分析済みのためスキップ`);
        return;
    }

    // 分析を実行
    const result = summarizePage(url);
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME).getRange(lineNumber, 8).setValue(result);
}


function testSummarizePage() {
    const url = "https://prtimes.jp/main/html/rd/p/000000014.000065715.html";
    const result = summarizePage(url);
    Logger.log(result);
}

// 指定されたURLページを読み、解説（買い手の戦略、目的）300文字くらいでまとめる
function summarizePage(url) {
    const response = UrlFetchApp.fetch(url);
    const htmlText = response.getContentText();
    // テキストだけを抽出
    const siteText = extractBodyText(htmlText);
    //// base64エンコード
    //const base64Html = Utilities.base64Encode(siteText);
    

    // htmlの内容をAIに渡して、解説（買い手の戦略、目的）300文字くらいでまとめる
    const payload = {
        model: CHAT_MODEL,
        messages: [
            {
                role: "system",
                content: [
                    "あなたは、M&Aや事業提携に関するビジネス戦略アナリストです。",
                    "与えられたプレスリリース等の公開情報から、買い手企業の戦略や目的を分析・要約する役割を担います。",
                    "抽出すべきのは、事実ではなく、背後にある企業の狙いや意図、経営上の戦略的意味です。",
                    "記載がない内容は推測せず、公開情報に基づいて慎重に整理してください。",
                    "出力は以下の形式に従ってください：",
                    "- 提携の背景と目的：",
                    "- 買い手企業の狙い：",
                    "- 提供する価値：",
                    "- 長期的な展望：",
                    "各項目は箇条書きでも構いません。300文字程度でわかりやすく簡潔にまとめてください。",
                ].join("\n"),
            },
            {
                role: "user",
                content: [
                    "以下の記事のデータを、system指示に従って分析してください。\n",
                    "\n",
                    "# 記事のデータ\n" + siteText + "\n",
                ].join("\n"),
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
    const responseAi = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);

    // レスポンスを取得して返す
    const result = JSON.parse(responseAi.getContentText());
    return result.choices[0].message.content;
}


/**
 * HTMLから<body>内のテキストを抽出して整形する
 * @param {string} html - 取得したHTML全体
 * @returns {string} - 抽出された本文テキスト
 */
function extractBodyText(html) {
    // bodyタグの中身だけを抽出（非貪欲マッチ）
    const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    if (!bodyMatch) {
        // bodyタグが見つからない場合は全体を対象にする
        return cleanHtml(html);
    }
  
    const bodyContent = bodyMatch[1];
    return cleanHtml(bodyContent);
}
  
  /**
   * HTMLタグを除去して、本文テキストだけを抽出する
   * @param {string} htmlFragment - HTML断片（body内）
   * @returns {string}
   */
function cleanHtml(htmlFragment) {
    return htmlFragment
        // scriptタグと中身を削除
        .replace(/<script[\s\S]*?<\/script>/gi, '')
        // styleタグと中身を削除
        .replace(/<style[\s\S]*?<\/style>/gi, '')
        // すべてのHTMLタグを除去
        .replace(/<\/?[^>]+>/g, '')
        // HTMLエンティティの基本的な置換（簡易）
        .replace(/&nbsp;/g, ' ')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&')
        // 連続する空白・改行を整形
        .replace(/\s+/g, ' ')
        .trim();
}