/**
 * @fileoverview
 * Google Drive上のMarkdownファイルから変更履歴を読み込み、
 * Webページとして表示するためのサーバーサイドスクリプトです。
 */

/**
 * WebアプリケーションにGETリクエストがあった場合に実行されるメイン関数です。
 * @param {Object} e イベントオブジェクト。
 * @return {HtmlOutput} 生成されたHTMLページ。
 */
function doGet(e) {
  try {
    // Google DriveからMarkdownファイルの内容を取得します。
    // MarkdownファイルのURLからIDを抽出して使用します。
    const fileId = "1srsKUw6dOJden8Dc3hBIDySppckc3d8t";
    const file = DriveApp.getFileById(fileId);
    const markdownText = file.getBlob().getDataAsString("UTF-8");

    // Markdownテキストを解析して、必要な情報を抽出します。
    const { title, description, history } = parseChangeLogMarkdown(markdownText);

    // history内のchangesプロパティに含まれるHTMLエンティティをデコードします。
    // これにより、テンプレート側で正しくHTMLとしてレンダリングされるようになります。
    history.forEach(item => {
      item.changes = unescapeHtml(item.changes);
    });

    // HTMLテンプレートを作成し、抽出したデータを渡します。
    const template = HtmlService.createTemplateFromFile('index');
    template.title = title;
    template.description = description;
    template.history = history;

    // テンプレートを評価してHTMLを生成し、クライアントに返します。
    // Webページのタイトルやビューポートも設定します。
    return template.evaluate()
      .setTitle(title || 'システムの変更履歴') // titleが空の場合のフォールバック
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  } catch (error) {
    // エラーが発生した場合は、エラーメッセージを含む単純なHTMLを返します。
    Logger.log(error.toString());
    return HtmlService.createHtmlOutput(
      `<p>エラーが発生しました。詳細はサーバーログを確認してください。</p><p>${error.toString()}</p>`
    );
  }
}

/**
 * 変更履歴のMarkdownテキストを解析し、タイトル、説明、履歴データのオブジェクトを返します。
 * @param {string} text 解析対象のMarkdownテキスト。
 * @return {{title: string, description: string, history: Array<Object>}} 解析結果。
 */
function parseChangeLogMarkdown(text) {
  const lines = text.split('\n');

  // 1. タイトルを取得
  const titleLine = lines.find(line => line.startsWith('# ')) || '';
  const title = titleLine.replace('# ', '').trim();
  const titleIndex = lines.findIndex(line => line.startsWith('# '));

  // 2. テーブルのヘッダーとセパレーターの位置を特定
  // 正規表現を使い、より柔軟にテーブルの開始行を判定
  const tableHeaderIndex = lines.findIndex(line => /\|.*バージョン.*\|/.test(line));
  const tableSeparatorIndex = lines.findIndex(line => /^\s*\|\s*:?-+:?\s*\|/.test(line));

  // 3. 説明文を取得
  let description = '';
  // タイトルとテーブルヘッダーの両方が見つかった場合のみ、その間の行を説明文とする
  if (titleIndex !== -1 && tableHeaderIndex !== -1 && titleIndex < tableHeaderIndex) {
    const descriptionLines = lines.slice(titleIndex + 1, tableHeaderIndex)
                                  .filter(line => line.trim() !== '');
    description = descriptionLines.join('<br>');
  }

  // 4. テーブルデータを抽出
  let history = [];
  // テーブルのセパレーターが見つかった場合のみ、それ以降の行を解析
  if (tableSeparatorIndex !== -1) {
    const tableDataLines = lines.slice(tableSeparatorIndex + 1)
                                  .filter(line => line.trim().startsWith('|'));

    history = tableDataLines.map(line => {
      const columns = line.split('|')
                          .map(col => col.trim())
                          .slice(1, -1);
      
      if (columns.length === 4) {
        return {
          version: columns[0].replace(/\*\*/g, ''),
          date: columns[1],
          author: columns[2],
          changes: convertChangesToHtml(columns[3]) // テキストをHTMLリストに変換
        };
      } else {
        // 解析に失敗した行をログに出力
        Logger.log(`警告: テーブル行の解析に失敗しました。列数が4ではありません。行: ${line}`);
        return null;
      }
    }).filter(item => item !== null);
  } else {
    Logger.log('警告: Markdownからテーブルの区切り行が見つかりませんでした。');
  }

  return { title, description, history };
}

/**
 * 変更内容のテキストをHTMLリストに変換します。
 * テキスト内の`<br>`タグをリスト項目(`<li>`)の区切りとして解釈します。
 * @param {string} text 変換するテキスト。例: "項目1<br>項目2"
 * @return {string} 変換されたHTMLリスト。例: "<ul><li>項目1</li><li>項目2</li></ul>"
 */
function convertChangesToHtml(text) {
  const trimmedText = text ? text.trim() : '';
  if (trimmedText === '') {
    return '';
  }

  // <br>, <br/>, <br /> などのタグで分割し、各項目をトリムして空の項目を除外
  const items = trimmedText.split(/<br\s*\/?>/i)
                           .map(item => item.trim())
                           .filter(item => item);

  // 単一項目でも必ず<ul><li>...</li></ul>で返す
  const listItems = items.map(item => `<li>${item}</li>`).join('');
  return `<ul>${listItems}</ul>`;
}

/**
 * HTMLエンティティを含む文字列をデコード（アンエスケープ）します。
 * 例: "&lt;p&gt;" を "<p>" に変換します。
 * @param {string} str デコードする文字列。
 * @return {string} デコードされた文字列。
 */
function unescapeHtml(str) {
  if (!str) return '';
  return str.replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&quot;/g, '"')
            .replace(/&#039;/g, "'")
            .replace(/&amp;/g, '&'); // &amp; は最後に置換することが重要です
}

/**
 * Markdownパーサーをデバッグするためのテスト関数です。
 * この関数をApps Scriptエディタから直接実行することで、
 * Webアプリをデプロイせずに解析結果をログで確認できます。
 */
function testParseChangeLogMarkdown() {
  try {
    // doGetと同じファイルIDを使用
    const fileId = "1srsKUw6dOJden8Dc3hBIDySppckc3d8t";
    const file = DriveApp.getFileById(fileId);
    const markdownText = file.getBlob().getDataAsString("UTF-8");

    // 実際に解析を実行します
    const parsedData = parseChangeLogMarkdown(markdownText);

    // 解析結果全体をログに出力します
    Logger.log("--- 解析結果全体 ---");
    // JSON.stringifyを使って、オブジェクトを人間が読みやすい形に整形して出力します
    Logger.log(JSON.stringify(parsedData, null, 2));

    // 各履歴の「変更内容」フィールドが取得できているか、より詳細にチェックします
    if (parsedData.history && parsedData.history.length > 0) {
      Logger.log("\n--- 各履歴の「変更内容」フィールドの個別チェック ---");
      parsedData.history.forEach((item, index) => {
        Logger.log(`[${index}] バージョン: ${item.version}`);
        Logger.log(`  変更内容 (changes): "${item.changes}"`); // 取得した文字列を""で囲んで表示
        Logger.log(`  取得文字数: ${item.changes ? item.changes.length : 0}`);
        Logger.log('--------------------');
      });
    }
  } catch (error) {
    Logger.log(`テスト実行中にエラーが発生しました: ${error.toString()}\n${error.stack}`);
  }
}
