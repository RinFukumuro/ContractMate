// 📌 弁護士業務支援用 Word文書ナビゲーションツール
// TypeScript (WSH + ActiveXObject)

declare var ActiveXObject: any;
declare var WScript: any;

class WordNavigator {
  private wordApp: any;
  private document: any;

  constructor() {
    try {
      this.wordApp = new ActiveXObject('Word.Application');
      this.wordApp.Visible = true;
    } catch (error) {
      WScript.Echo('Wordの起動に失敗しました。: ' + error);
      WScript.Quit();
    }
  }

  // Word文書を開く
  openDocument(filePath: string): boolean {
    try {
      const fso = new ActiveXObject('Scripting.FileSystemObject');
      const absPath = fso.GetAbsolutePathName(filePath);
      this.document = this.wordApp.Documents.Open(absPath);
      return true;
    } catch (error) {
      WScript.Echo('文書を開くことができませんでした。: ' + error);
      return false;
    }
  }

  // 文書の段落構造を取得（条項や段落を識別するため）
  getParagraphStructure(): Array<any> {
    if (!this.document) {
      WScript.Echo('文書が開かれていません。');
      return [];
    }

    const structure = [];
    const paragraphCount = this.document.Paragraphs.Count;

    for (let i = 1; i <= paragraphCount; i++) {
      const para = this.document.Paragraphs(i);
      const paraText = para.Range.Text.trim();
      const style = para.Range.Style.NameLocal;
      const outlineLevel = para.OutlineLevel;

      structure.push({
        index: i,
        text: paraText,
        style: style,
        outlineLevel: outlineLevel,
      });
    }
    return structure;
  }

  // 指定した段落番号に即座に移動
  navigateToParagraph(paragraphNumber: number): boolean {
    if (!this.document) {
      WScript.Echo('文書が開かれていません。');
      return false;
    }

    try {
      const totalParagraphs = this.document.Paragraphs.Count;
      if (paragraphNumber < 1 || paragraphNumber > totalParagraphs) {
        WScript.Echo(`段落番号が無効です: ${paragraphNumber}`);
        return false;
      }

      const para = this.document.Paragraphs(paragraphNumber);
      para.Range.Select();
      this.wordApp.Activate();
      return true;
    } catch (error) {
      WScript.Echo('指定した段落に移動できませんでした。: ' + error);
      return false;
    }
  }

  // 文書を閉じる
  closeDocument(save: boolean = false): void {
    if (this.document) {
      if (save) {
        this.document.Save();
      }
      this.document.Close(save);
      this.document = null;
    }
  }

  // Wordアプリケーションを終了
  quit(): void {
    if (this.wordApp) {
      this.wordApp.Quit();
      this.wordApp = null;
    }
  }
}

// メイン処理
function main() {
  const args = WScript.Arguments;
  if (args.length < 2) {
    WScript.Echo('使用方法: cscript wordNavigator.js [Word文書パス] [移動したい段落番号]');
    WScript.Quit();
  }

  const filePath = args(0);
  const targetParagraph = parseInt(args(1), 10);

  const navigator = new WordNavigator();

  if (navigator.openDocument(filePath)) {
    const structure = navigator.getParagraphStructure();

    // 文書構造を表示（必要に応じて確認可能）
    for (const para of structure) {
      WScript.Echo(`段落番号: ${para.index}, スタイル: ${para.style}, アウトラインレベル: ${para.outlineLevel}, 内容: ${para.text}`);
    }

    // 指定段落へ即座に移動
    if (!navigator.navigateToParagraph(targetParagraph)) {
      WScript.Echo('指定した段落への移動に失敗しました。');
    }
  }
}

// スクリプト実行
main();