// =========================================================================
// Word.Applicationの基本テストスクリプト
// Windows Script Host (WSH) + ActiveXObject
// =========================================================================

try {
    WScript.Echo("===== Word基本テスト開始 =====");
    WScript.Echo("実行日時: " + new Date().toLocaleString());
    
    // FileSystemObjectのテスト
    WScript.Echo("FileSystemObjectのテスト...");
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        WScript.Echo("✓ FileSystemObjectの作成に成功しました");
    } catch (e) {
        WScript.Echo("✗ FileSystemObjectの作成に失敗しました: " + e.message);
        WScript.Quit(1);
    }
    
    // Word.Applicationのテスト
    WScript.Echo("Word.Applicationのテスト...");
    try {
        var wordApp = new ActiveXObject("Word.Application");
        WScript.Echo("✓ Word.Applicationの作成に成功しました");
        
        // Wordのバージョン情報
        WScript.Echo("Wordのバージョン: " + wordApp.Version);
        
        // Wordを表示
        wordApp.Visible = true;
        WScript.Echo("✓ Wordを表示しました");
        
        // 新しい文書を作成
        WScript.Echo("新しい文書を作成しています...");
        var doc = wordApp.Documents.Add();
        WScript.Echo("✓ 新しい文書を作成しました");
        
        // テキストを追加
        WScript.Echo("テキストを追加しています...");
        var selection = wordApp.Selection;
        selection.TypeText("これはテストです。Word.Applicationが正常に動作しています。");
        selection.TypeParagraph();
        selection.TypeText("このテストが成功すれば、WordNavigatorも動作するはずです。");
        WScript.Echo("✓ テキストを追加しました");
        
        // 3秒待機
        WScript.Echo("3秒間待機します...");
        WScript.Sleep(3000);
        
        // 文書を閉じる（保存しない）
        WScript.Echo("文書を閉じています...");
        doc.Close(false);
        WScript.Echo("✓ 文書を閉じました");
        
        // Wordを終了
        WScript.Echo("Wordを終了しています...");
        wordApp.Quit();
        WScript.Echo("✓ Wordを終了しました");
    } catch (e) {
        WScript.Echo("✗ Wordの操作中にエラーが発生しました: " + e.message);
        
        // Wordを強制終了
        try {
            if (wordApp) {
                wordApp.Quit();
            }
        } catch (quitError) {
            // 無視
        }
        
        WScript.Quit(1);
    }
    
    WScript.Echo("===== テスト成功 =====");
} catch (e) {
    WScript.Echo("✗ 致命的なエラーが発生しました: " + e.message);
    WScript.Quit(1);
}