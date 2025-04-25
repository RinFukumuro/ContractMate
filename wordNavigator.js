// 📌 弁護士業務支援用 Word文書ナビゲーションツール
// Windows Script Host (WSH) + ActiveXObject

var WordNavigator = /** @class */ (function () {
    function WordNavigator() {
        try {
            this.wordApp = new ActiveXObject('Word.Application');
            this.wordApp.Visible = true;
        }
        catch (error) {
            WScript.Echo('Wordの起動に失敗しました。: ' + error);
            WScript.Quit();
        }
    }
    // Word文書を開く
    WordNavigator.prototype.openDocument = function (filePath) {
        try {
            var fso = new ActiveXObject('Scripting.FileSystemObject');
            var absPath = fso.GetAbsolutePathName(filePath);
            this.document = this.wordApp.Documents.Open(absPath);
            return true;
        }
        catch (error) {
            WScript.Echo('文書を開くことができませんでした。: ' + error);
            return false;
        }
    };
    // 文書の段落構造を取得（条項や段落を識別するため）
    WordNavigator.prototype.getParagraphStructure = function () {
        if (!this.document) {
            WScript.Echo('文書が開かれていません。');
            return [];
        }
        var structure = [];
        var paragraphCount = this.document.Paragraphs.Count;
        for (var i = 1; i <= paragraphCount; i++) {
            var para = this.document.Paragraphs(i);
            var paraText = para.Range.Text.trim();
            var style = para.Range.Style.NameLocal;
            var outlineLevel = para.OutlineLevel;
            structure.push({
                index: i,
                text: paraText,
                style: style,
                outlineLevel: outlineLevel
            });
        }
        return structure;
    };
    // 指定した段落番号に即座に移動
    WordNavigator.prototype.navigateToParagraph = function (paragraphNumber) {
        if (!this.document) {
            WScript.Echo('文書が開かれていません。');
            return false;
        }
        try {
            var totalParagraphs = this.document.Paragraphs.Count;
            if (paragraphNumber < 1 || paragraphNumber > totalParagraphs) {
                WScript.Echo("\u6BB5\u843D\u756A\u53F7\u304C\u7121\u52B9\u3067\u3059: ".concat(paragraphNumber));
                return false;
            }
            var para = this.document.Paragraphs(paragraphNumber);
            para.Range.Select();
            this.wordApp.Activate();
            return true;
        }
        catch (error) {
            WScript.Echo('指定した段落に移動できませんでした。: ' + error);
            return false;
        }
    };
    // 文書を閉じる
    WordNavigator.prototype.closeDocument = function (save) {
        if (save === void 0) { save = false; }
        if (this.document) {
            if (save) {
                this.document.Save();
            }
            this.document.Close(save);
            this.document = null;
        }
    };
    // Wordアプリケーションを終了
    WordNavigator.prototype.quit = function () {
        if (this.wordApp) {
            this.wordApp.Quit();
            this.wordApp = null;
        }
    };
    return WordNavigator;
}());

// メイン処理
function main() {
    var args = WScript.Arguments;
    if (args.length < 2) {
        WScript.Echo('使用方法: cscript wordNavigator.js [Word文書パス] [移動したい段落番号]');
        WScript.Quit();
    }
    var filePath = args(0);
    var targetParagraph = parseInt(args(1), 10);
    var navigator = new WordNavigator();
    if (navigator.openDocument(filePath)) {
        var structure = navigator.getParagraphStructure();
        // 文書構造を表示（必要に応じて確認可能）
        for (var i = 0; i < structure.length; i++) {
            var para = structure[i];
            WScript.Echo("\u6BB5\u843D\u756A\u53F7: ".concat(para.index, ", \u30B9\u30BF\u30A4\u30EB: ").concat(para.style, ", \u30A2\u30A6\u30C8\u30E9\u30A4\u30F3\u30EC\u30D9\u30EB: ").concat(para.outlineLevel, ", \u5185\u5BB9: ").concat(para.text));
        }
        // 指定段落へ即座に移動
        if (!navigator.navigateToParagraph(targetParagraph)) {
            WScript.Echo('指定した段落への移動に失敗しました。');
        }
    }
}

// スクリプト実行
main();