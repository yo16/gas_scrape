// M&A速報(https://www.marr.jp/genre/topics/news/)をスクレイプする
function scrapeMarrTopicsNews() {
    const url = 'https://www.marr.jp/genre/topics/news/';
    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();
  
    // <ul class="textCassetteWrapper"> ～ </ul> を抜き出す
    const ulMatch = html.match(
        /<ul[^>]*class="[^"]*textCassetteWrapper[^"]*"[^>]*>([\s\S]*?)<\/ul>/
    );
  
    if (!ulMatch) {
        Logger.log('対象の<ul class="textCassetteWrapper">が見つかりません');
        return;
    }
  
    // ULの中身だけ抜き出す
    const ulContent = ulMatch[1];
  
    // <li>要素を配列にして取り出す
    // <li id="61243" class="textCassette">～</li>
    const liMatches = [...ulContent.matchAll(/<li[^>]*>[\s\S]*?<\/li>/g)];
  
    const results = liMatches.map(match => {
        return parseTopicsNewsItem(match[0]);
    });
  
    //Logger.log(results);
    // シートに出力
    writeToSheet(results, "M&A速報");
}


// M&A速報(https://www.marr.jp/genre/topics/news/)の<li>要素をパースする
function parseTopicsNewsItem(liElementStr) {
    // liElementStrは以下のような形式
    // <li id="61276" class="textCassette">
    //     <a class="textCassette__title textUnderline " href="/genre/topics/news/entry/61276">ヒューリック&lt;3003&gt;、鉱研工業&lt;6297&gt;に対しTOBを実施 買付価格は1株764円 同社は「賛同」を表明</a>
    //     <div class="textCassette__column">
    //         <p class="textCassette__subGenre">
    //         [M&amp;A速報]
    //                                 2025年06月17日(火)
    //                                                     </p>
    //     </div>
    //     <p class="textCassette__author"></p>
    // </li>

    // liElementStrから<li>の中身を取り出す
    const liElementMatch = liElementStr.match(/(<li[^>]*>)([\s\S]*?)<\/li>/);
    const liElement = liElementMatch ? liElementMatch[1] : '';
    const liInnerHtml = liElementMatch ? liElementMatch[2] : '';

    // idを取り出す
    const idMatch = liElement.match(/id="([^"]*)"/);
    const id = idMatch ? idMatch[1].trim() : '';

    // 記事の中身
    // タイトル
    const titleMatch = liInnerHtml.match(/<a[^>]*>(.*?)<\/a>/);
    const title = titleMatch ? titleMatch[1].trim() : '';
    // ページリンク
    const urlMatch = liInnerHtml.match(/<a [\s\S]*?href="(.*?)"/);
    const url = urlMatch ? 'https://www.marr.jp' + urlMatch[1] : '';
    // ジャンルと日付
    const genreAndDateMatch = liInnerHtml.match(
        /<p[^>]*class="[^"]*textCassette__subGenre[^"]*"[^>]*>\s*\[([^\]]+)\]\s*([\d年月日火水木金土\(\)]+)/
    );
      
    const genre = genreAndDateMatch ? genreAndDateMatch[1].trim() : '';
    const date = genreAndDateMatch ? genreAndDateMatch[2].trim() : '';

    return {
        id: parseInt(id),
        title: unescapeHtml(title),
        url,
        genre: unescapeHtml(genre),
        date,
    };
}


// シートへ出力
const writeToSheet = (datas, sheetTitle) => {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTitle);
    if (!sheet) {
        sheet = setServers.insertSheet(sheetTitle);
        sheet.appendRow(["id", "title", "url", "genre", "date"]);
    }

    // シートの最終行を取得
    const lastRow = sheet.getLastRow();
    // シートの最終列を取得
    const lastColumn = sheet.getLastColumn();

    // 1列目を取得
    const column1 = sheet.getRange(1, 1, lastRow, 1);  // ヘッダーを除く
    const column1Data = column1.getValues().flat();

    // １要素ずつ、行末に追加する
    // ただし、1列目だけを検索して、idがすでに存在する場合はスキップする
    datas.forEach(data => {
        if (!column1Data.includes(data.id)) {
            sheet.appendRow([data.id, data.title, data.url, data.genre, data.date]);
        }
    });
    
    // シートの最終行を取得
    const lastRowUpdated = sheet.getLastRow();

    // 1列目を降順でソートする
    sheet.getRange(2, 1, lastRowUpdated - 1, lastColumn).sort({column: 1, ascending: false});
}


// htmlエスケープを戻す
const unescapeHtml = (str) => {
    return str.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
