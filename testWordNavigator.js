// WordNavigatorのテスト用スクリプト

// ファイルシステムオブジェクトを作成
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");
var currentDir = fso.GetAbsolutePathName(".");

// サンプル契約書のパス
var sampleDocPath = currentDir + "\\サンプル契約書.docx";

// コマンドライン引数を取得
var args = WScript.Arguments;
var targetParagraph = 1; // デフォルトは1段落目

if (args.length > 0) {
    targetParagraph = parseInt(args(0), 10);
}

// サンプル契約書が存在するか確認
if (!fso.FileExists(sampleDocPath)) {
    WScript.Echo("サンプル契約書が見つかりません。自動生成します...");
    
    try {
        // createSampleDoc.jsを実行
        shell.Run("cscript " + currentDir + "\\createSampleDoc.js", 1, true);
        
        // 再度ファイルの存在を確認
        if (!fso.FileExists(sampleDocPath)) {
            WScript.Echo("サンプル契約書の生成に失敗しました。");
            WScript.Quit(1);
        }
    } catch (error) {
        WScript.Echo("サンプル契約書の生成中にエラーが発生しました: " + error);
        WScript.Quit(1);
    }
}

WScript.Echo("WordNavigatorを使用してサンプル契約書を開きます...");
WScript.Echo("指定された段落番号: " + targetParagraph);

// wordNavigator.jsを実行
try {
    var command = "cscript " + currentDir + "\\wordNavigator.js \"" + sampleDocPath + "\" " + targetParagraph;
    WScript.Echo("実行コマンド: " + command);
    shell.Run(command, 1, true);
} catch (error) {
    WScript.Echo("WordNavigatorの実行中にエラーが発生しました: " + error);
    WScript.Quit(1);
}

// テスト用のヘルプメッセージ
WScript.Echo("\n=== WordNavigatorテストヘルプ ===");
WScript.Echo("以下のコマンドで特定の段落に移動できます：");
WScript.Echo("cscript testWordNavigator.js 5  # 第1条（目的）に移動");
WScript.Echo("cscript testWordNavigator.js 13 # 第3条（契約期間）に移動");
WScript.Echo("cscript testWordNavigator.js 19 # 第5条（機密保持）に移動");