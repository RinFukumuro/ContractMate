// WordNavigatorのテスト用スクリプト（simple_navigator.js版）

// ファイルシステムオブジェクトを作成
var fso = new ActiveXObject("Scripting.FileSystemObject");
var shell = new ActiveXObject("WScript.Shell");

// サンプル契約書のパス
var sampleDocPath = "C:\\Users\\Orinr\\Documents\\sample_contract.docx";

// コマンドライン引数を取得
var args = WScript.Arguments;
var targetParagraph = 1; // デフォルトは1段落目

if (args.length > 0) {
    targetParagraph = parseInt(args(0), 10);
}

// サンプル契約書が存在するか確認
if (!fso.FileExists(sampleDocPath)) {
    WScript.Echo("指定されたサンプル契約書が見つかりません: " + sampleDocPath);
    WScript.Quit(1);
}

WScript.Echo("Simple Navigatorを使用してサンプル契約書を開きます...");
WScript.Echo("指定された段落番号: " + targetParagraph);

// simple_navigator.jsを実行
try {
    var currentDir = fso.GetAbsolutePathName(".");
    var scriptPath = "C:\\Users\\Orinr\\OneDrive\\ドキュメント\\Legal Agent\\3日目\\VScodeExtension1\\ContractMate6\\simple_navigator.js";
    var command = "cscript \"" + scriptPath + "\" \"" + sampleDocPath + "\" " + targetParagraph;
    WScript.Echo("実行コマンド: " + command);
    shell.Run(command, 1, true);
} catch (error) {
    WScript.Echo("Simple Navigatorの実行中にエラーが発生しました: " + error);
    WScript.Quit(1);
}

// テスト用のヘルプメッセージ
WScript.Echo("\n=== WordNavigatorテストヘルプ ===");
WScript.Echo("以下のコマンドで特定の段落に移動できます：");
WScript.Echo("cscript testSimpleNavigator.js 5  # 第1条（目的）に移動");
WScript.Echo("cscript testSimpleNavigator.js 13 # 第3条（契約期間）に移動");
WScript.Echo("cscript testSimpleNavigator.js 19 # 第5条（機密保持）に移動");