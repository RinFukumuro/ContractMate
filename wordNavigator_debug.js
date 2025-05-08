// =========================================================================
// 弁護士業務支援用 Word文書ナビゲーションツール（デバッグ版）
// Windows Script Host (WSH) + ActiveXObject
// =========================================================================

// デバッグ情報の表示
try {
    WScript.Echo("===== デバッグ情報 =====");
    WScript.Echo("スクリプト実行開始: " + new Date().toLocaleString());
    WScript.Echo("Windows Script Host バージョン: " + WScript.Version);
    
    // 引数の確認
    WScript.Echo("引数の数: " + WScript.Arguments.length);
    for (var i = 0; i < WScript.Arguments.length; i++) {
        WScript.Echo("引数" + i + ": " + WScript.Arguments(i));
    }
    
    // FileSystemObjectのテスト
    WScript.Echo("FileSystemObjectのテスト...");
    try {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        WScript.Echo("✓ FileSystemObjectの作成に成功しました");
        
        // カレントディレクトリの表示
        var currentDir = fso.GetAbsolutePathName(".");
        WScript.Echo("カレントディレクトリ: " + currentDir);
        
        // ファイルの存在確認
        if (WScript.Arguments.length > 0) {
            var filePath = WScript.Arguments(0);
            WScript.Echo("ファイルパス: " + filePath);
            if (fso.FileExists(filePath)) {
                WScript.Echo("✓ ファイルは存在します: " + filePath);
                
                // ファイル情報の表示
                var file = fso.GetFile(filePath);
                WScript.Echo("ファイルサイズ: " + file.Size + " バイト");
                WScript.Echo("最終更新日時: " + file.DateLastModified);
            } else {
                WScript.Echo("✗ エラー: ファイルが存在しません: " + filePath);
                WScript.Quit(1);
            }
        }
    } catch (e) {
        WScript.Echo("✗ FileSystemObjectの作成に失敗しました: " + e.message);
        WScript.Echo("  原因: ActiveXObjectの使用が許可されていない可能性があります");
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
        
        // 文書を開く
        if (WScript.Arguments.length > 0) {
            var filePath = WScript.Arguments(0);
            try {
                WScript.Echo("文書を開いています: " + filePath);
                var absPath = fso.GetAbsolutePathName(filePath);
                var doc = wordApp.Documents.Open(absPath);
                WScript.Echo("✓ 文書を開きました: " + absPath);
                
                // 文書情報の表示
                WScript.Echo("文書名: " + doc.Name);
                WScript.Echo("段落数: " + doc.Paragraphs.Count);
                
                // 段落情報の表示（最初の5つ）
                WScript.Echo("最初の5つの段落:");
                var maxPara = Math.min(5, doc.Paragraphs.Count);
                for (var i = 1; i <= maxPara; i++) {
                    var para = doc.Paragraphs(i);
                    var paraText = para.Range.Text;
                    if (paraText) {
                        paraText = paraText.replace(/\r/g, "").replace(/\n/g, "");
                    } else {
                        paraText = "";
                    }
                    
                    // 長いテキストは省略
                    if (paraText.length > 50) {
                        paraText = paraText.substring(0, 50) + "...";
                    }
                    
                    WScript.Echo("段落" + i + ": " + paraText);
                }
                
                // 段落移動のテスト
                if (WScript.Arguments.length > 1 && WScript.Arguments(1) !== "-k" && WScript.Arguments(1) !== "-h") {
                    var targetParagraph = parseInt(WScript.Arguments(1), 10);
                    if (!isNaN(targetParagraph) && targetParagraph > 0 && targetParagraph <= doc.Paragraphs.Count) {
                        WScript.Echo("段落" + targetParagraph + "に移動します...");
                        var para = doc.Paragraphs(targetParagraph);
                        para.Range.Select();
                        wordApp.Activate();
                        WScript.Echo("✓ 段落に移動しました");
                    } else {
                        WScript.Echo("✗ 無効な段落番号です: " + targetParagraph);
                    }
                }
                
                // 文書は閉じない（ユーザーが確認できるようにするため）
                WScript.Echo("✓ 文書の処理が完了しました。確認後、手動でWordを閉じてください。");
            } catch (docError) {
                WScript.Echo("✗ 文書を開く際にエラーが発生しました: " + docError.message);
                
                // 詳細なエラー情報
                if (docError.number) {
                    WScript.Echo("  エラーコード: " + docError.number);
                }
                
                // Wordは終了する
                wordApp.Quit();
                WScript.Quit(1);
            }
        } else {
            // 引数がない場合はWordを終了
            wordApp.Quit();
        }
    } catch (e) {
        WScript.Echo("✗ Word.Applicationの作成に失敗しました: " + e.message);
        WScript.Echo("  原因: Microsoft Wordがインストールされていないか、ActiveXObjectの使用が許可されていない可能性があります");
        WScript.Quit(1);
    }
    
    WScript.Echo("===== デバッグ情報終了 =====");
} catch (e) {
    WScript.Echo("✗ 致命的なエラーが発生しました: " + e.message);
    WScript.Quit(1);
}

// スクリプト実行完了
WScript.Echo("スクリプト実行完了: " + new Date().toLocaleString());