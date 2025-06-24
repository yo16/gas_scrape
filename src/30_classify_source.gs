// リンク先を見て、htmlかpdfかを判断する

const TARGET_SHEET_NAME = "M&A速報";


// F列目がTRUEの行を探して、その行に対してsetSourceStringを実行する
function setSourceStringsIfTrue() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    for (let i = 2; i <= lastRow; i++) {
        const sourceString = sheet.getRange(i, 6).getValue();
        // TRUEであるときに、setSourceStringを実行する
        if (sourceString === true) {
            setSourceString(i, TARGET_SHEET_NAME);
        }
    }
}

function testSetSourceString() {
    setSourceString(25, TARGET_SHEET_NAME);
}
// シートのlineNumber行目のC列のURLを開き、G列にsourceStringを設定する
function setSourceString(lineNumber, sheetName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const url = sheet.getRange(lineNumber, 3).getValue();

    let sourceUrl = getSourceUrl(url);
    if (!sourceUrl) {
        return;
    }

    // もし`/`で始まっている場合は、`https://www.marr.jp`を付与する
    if (sourceUrl.startsWith("/")) {
        sourceUrl = "https://www.marr.jp" + sourceUrl;
    }

    sheet.getRange(lineNumber, 7).setValue(sourceUrl);
}


function testGetSourceUrl() {
    const sourceUrl = getSourceUrl("https://www.marr.jp/genre/topics/news/entry/61352");
    Logger.log(sourceUrl);
}
// MARRのトピックスページを読んで、そこに書かれているリンクを取得する
function getSourceUrl(marrSiteUrl) {
    const response = UrlFetchApp.fetch(marrSiteUrl);
    const html = response.getContentText();

    const sourceUrl = html.match(/<ul class="dairy">[\s\S]*?<a href="([^"]+)"/);
    if (!sourceUrl) {
        Logger.log("sourceUrl not found");
        return;
    }

    return sourceUrl[1];
}


function classifySource(url) {
}