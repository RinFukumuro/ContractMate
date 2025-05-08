// =========================================================================
// 弁護士業務支援用 Word文書ナビゲーションツール（統合版）
// Windows Script Host (WSH) + ActiveXObject
// =========================================================================

/**
 * WordNavigator クラス
 * Word文書を開き、段落構造を解析し、特定の段落に移動するための機能を提供します。
 */
var WordNavigator = (function () {
    /**
     * WordNavigatorのコンストラクタ
     * Word.Applicationを起動し、表示状態にします。
     */
    function WordNavigator() {
        try {
            WScript.Echo("Word.Applicationを起動しています...");
            this.wordApp = new ActiveXObject("Word.Application");
            this.wordApp.Visible = true;
            WScript.Echo("Word.Applicationの起動に成功しました");
        } catch (error) {
            WScript.Echo("Wordの起動に失敗しました。: " + error);
            WScript.Quit(1);
        }
    }
    
    /**
     * Word文書を開きます
     * @param {string} filePath - 開くWord文書のパス
     * @returns {boolean} - 成功した場合はtrue、失敗した場合はfalse
     */
    WordNavigator.prototype.openDocument = function (filePath) {
        try {
            WScript.Echo("文書を開いています: " + filePath);
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            
            // ファイルの存在確認
            if (!fso.FileExists(filePath)) {
                WScript.Echo("エラー: ファイルが存在しません - " + filePath);
                return false;
            }
            
            var absPath = fso.GetAbsolutePathName(filePath);
            this.document = this.wordApp.Documents.Open(absPath);
            WScript.Echo("文書を開きました: " + absPath);
            return true;
        } catch (error) {
            WScript.Echo("文書を開くことができませんでした。: " + error);
            return false;
        }
    };
    
    /**
     * 文書の段落構造を取得します（条項や段落を識別するため）
     * @returns {Array} - 段落情報の配列
     */
    WordNavigator.prototype.getParagraphStructure = function () {
        if (!this.document) {
            WScript.Echo("文書が開かれていません。");
            return [];
        }
        
        try {
            WScript.Echo("段落構造を解析しています...");
            var structure = [];
            var paragraphCount = this.document.Paragraphs.Count;
            WScript.Echo("段落数: " + paragraphCount);
            
            for (var i = 1; i <= paragraphCount; i++) {
                var para = this.document.Paragraphs(i);
                var paraText = para.Range.Text;
                
                // テキスト処理の改善（改行文字の削除など）
                if (paraText) {
                    paraText = paraText.replace(/\r/g, "").replace(/\n/g, "");
                } else {
                    paraText = "";
                }
                
                var style = para.Range.Style.NameLocal;
                var outlineLevel = para.OutlineLevel;
                
                structure.push({
                    index: i,
                    text: paraText,
                    style: style,
                    outlineLevel: outlineLevel
                });
                
                // 進捗表示（大きな文書の場合）
                if (paragraphCount > 100 && i % 50 === 0) {
                    WScript.Echo("解析中... " + Math.round((i / paragraphCount) * 100) + "%完了");
                }
            }
            
            WScript.Echo("段落構造の解析が完了しました");
            return structure;
        } catch (error) {
            WScript.Echo("段落構造の取得中にエラーが発生しました: " + error);
            return [];
        }
    };
    
    /**
     * 指定した段落番号に即座に移動します
     * @param {number} paragraphNumber - 移動したい段落番号
     * @returns {boolean} - 成功した場合はtrue、失敗した場合はfalse
     */
    WordNavigator.prototype.navigateToParagraph = function (paragraphNumber) {
        if (!this.document) {
            WScript.Echo("文書が開かれていません。");
            return false;
        }
        
        try {
            var totalParagraphs = this.document.Paragraphs.Count;
            if (paragraphNumber < 1 || paragraphNumber > totalParagraphs) {
                WScript.Echo("段落番号が無効です: " + paragraphNumber);
                WScript.Echo("有効な範囲: 1～" + totalParagraphs);
                return false;
            }
            
            WScript.Echo(paragraphNumber + "番目の段落に移動します...");
            var para = this.document.Paragraphs(paragraphNumber);
            para.Range.Select();
            this.wordApp.Activate();
            WScript.Echo("段落に移動しました");
            return true;
        } catch (error) {
            WScript.Echo("指定した段落に移動できませんでした。: " + error);
            return false;
        }
    };
    
    /**
     * 特定のキーワードを含む段落を検索して移動します
     * @param {string} keyword - 検索するキーワード
     * @returns {boolean} - 成功した場合はtrue、失敗した場合はfalse
     */
    WordNavigator.prototype.navigateToKeyword = function (keyword) {
        if (!this.document) {
            WScript.Echo("文書が開かれていません。");
            return false;
        }
        
        try {
            WScript.Echo("キーワード「" + keyword + "」を検索しています...");
            var found = false;
            var structure = this.getParagraphStructure();
            
            for (var i = 0; i < structure.length; i++) {
                if (structure[i].text.indexOf(keyword) >= 0) {
                    WScript.Echo("キーワードが見つかりました: 段落番号 " + structure[i].index);
                    this.navigateToParagraph(structure[i].index);
                    found = true;
                    break;
                }
            }
            
            if (!found) {
                WScript.Echo("キーワード「" + keyword + "」は文書内に見つかりませんでした。");
            }
            
            return found;
        } catch (error) {
            WScript.Echo("キーワード検索中にエラーが発生しました: " + error);
            return false;
        }
    };
    
    /**
     * 見出しスタイルの段落のみを抽出します（目次生成などに利用可能）
     * @returns {Array} - 見出し段落の配列
     */
    WordNavigator.prototype.getHeadings = function () {
        if (!this.document) {
            WScript.Echo("文書が開かれていません。");
            return [];
        }
        
        try {
            WScript.Echo("見出しを抽出しています...");
            var structure = this.getParagraphStructure();
            var headings = [];
            
            for (var i = 0; i < structure.length; i++) {
                var para = structure[i];
                // 見出しスタイルまたはアウトラインレベルが9未満の段落を見出しとして扱う
                if (para.style.indexOf("見出し") >= 0 || para.outlineLevel < 9) {
                    headings.push(para);
                }
            }
            
            WScript.Echo(headings.length + "個の見出しが見つかりました");
            return headings;
        } catch (error) {
            WScript.Echo("見出し抽出中にエラーが発生しました: " + error);
            return [];
        }
    };
    
    /**
     * 文書を閉じます
     * @param {boolean} save - 変更を保存する場合はtrue、破棄する場合はfalse
     */
    WordNavigator.prototype.closeDocument = function (save) {
        if (save === undefined) { save = false; }
        
        if (this.document) {
            try {
                WScript.Echo("文書を閉じています" + (save ? "（保存します）" : "（保存しません）"));
                if (save) {
                    this.document.Save();
                }
                this.document.Close(save);
                this.document = null;
                WScript.Echo("文書を閉じました");
            } catch (error) {
                WScript.Echo("文書を閉じる際にエラーが発生しました: " + error);
            }
        }
    };
    
    /**
     * Wordアプリケーションを終了します
     */
    WordNavigator.prototype.quit = function () {
        if (this.wordApp) {
            try {
                WScript.Echo("Wordアプリケーションを終了しています...");
                this.wordApp.Quit();
                this.wordApp = null;
                WScript.Echo("Wordアプリケーションを終了しました");
            } catch (error) {
                WScript.Echo("Wordアプリケーションの終了中にエラーが発生しました: " + error);
            }
        }
    };
    
    return WordNavigator;
})();

/**
 * メイン処理
 * コマンドライン引数を解析し、WordNavigatorを実行します
 */
function main() {
    try {
        WScript.Echo("WordNavigator統合版を実行しています...");
        var args = WScript.Arguments;
        
        // コマンドライン引数の解析
        if (args.length < 1) {
            WScript.Echo("使用方法:");
            WScript.Echo("  基本: cscript wordNavigator_integrated.js [Word文書パス] [移動したい段落番号]");
            WScript.Echo("  キーワード検索: cscript wordNavigator_integrated.js [Word文書パス] -k [検索キーワード]");
            WScript.Echo("  見出し表示: cscript wordNavigator_integrated.js [Word文書パス] -h");
            WScript.Quit(0);
        }
        
        var filePath = args(0);
        var navigator = new WordNavigator();
        
        if (navigator.openDocument(filePath)) {
            // 文書が正常に開かれた場合の処理
            
            if (args.length > 1) {
                var secondArg = args(1);
                
                // キーワード検索モード
                if (secondArg === "-k" && args.length > 2) {
                    var keyword = args(2);
                    navigator.navigateToKeyword(keyword);
                }
                // 見出し表示モード
                else if (secondArg === "-h") {
                    var headings = navigator.getHeadings();
                    WScript.Echo("===== 文書の見出し構造 =====");
                    for (var i = 0; i < headings.length; i++) {
                        var heading = headings[i];
                        var indent = "";
                        for (var j = 0; j < heading.outlineLevel; j++) {
                            indent += "  ";
                        }
                        WScript.Echo(indent + heading.index + ": " + heading.text);
                    }
                }
                // 段落番号指定モード
                else {
                    var targetParagraph = parseInt(secondArg, 10);
                    if (isNaN(targetParagraph)) {
                        WScript.Echo("エラー: 無効な段落番号です - " + secondArg);
                    } else {
                        // 文書構造を取得して表示
                        var structure = navigator.getParagraphStructure();
                        
                        // 文書構造を表示（必要に応じて確認可能）
                        WScript.Echo("===== 文書構造 =====");
                        for (var i = 0; i < structure.length; i++) {
                            var para = structure[i];
                            WScript.Echo("段落番号: " + para.index + ", スタイル: " + para.style + 
                                      ", アウトラインレベル: " + para.outlineLevel + ", 内容: " + para.text);
                        }
                        
                        // 指定段落へ即座に移動
                        if (!navigator.navigateToParagraph(targetParagraph)) {
                            WScript.Echo("指定した段落への移動に失敗しました。");
                        }
                    }
                }
            } else {
                // 引数が1つだけの場合は文書構造のみを表示
                var structure = navigator.getParagraphStructure();
                WScript.Echo("===== 文書構造 =====");
                for (var i = 0; i < structure.length; i++) {
                    var para = structure[i];
                    WScript.Echo("段落番号: " + para.index + ", スタイル: " + para.style + 
                              ", アウトラインレベル: " + para.outlineLevel + ", 内容: " + para.text);
                }
            }
            
            // 注意: 自動的に文書を閉じたりWordを終了したりしないようにしています
            // ユーザーが文書を確認した後、手動でWordを閉じることを想定しています
            WScript.Echo("処理が完了しました。確認後、Wordを手動で閉じてください。");
        }
    } catch (error) {
        WScript.Echo("エラーが発生しました: " + error);
        WScript.Quit(1);
    }
}

// スクリプト実行
main();