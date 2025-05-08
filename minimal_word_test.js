// 最小限のWord文書テストスクリプト

try {
    // Wordアプリケーションを起動
    var word = new ActiveXObject("Word.Application");
    word.Visible = true;
    
    // ファイルパスを指定
    var filePath = "C:\\Users\\Orinr\\Documents\\sample_contract.docx";
    
    // ファイルを開く
    var doc = word.Documents.Open(filePath);
    
    // 文書の段落数を表示
    var paragraphCount = doc.Paragraphs.Count;
    WScript.Echo("文書の段落数: " + paragraphCount);
    
    // 最初の5つの段落のテキストを表示
    WScript.Echo("\n最初の5つの段落:");
    var maxParagraphs = Math.min(5, paragraphCount);
    for (var i = 1; i <= maxParagraphs; i++) {
        var para = doc.Paragraphs(i);
        WScript.Echo(i + ": " + para.Range.Text);
    }
    
    // 文書を閉じる（保存しない）
    doc.Close(false);
    
    // Wordを終了
    word.Quit();
    
} catch (error) {
    WScript.Echo("エラーが発生しました: " + error);
}