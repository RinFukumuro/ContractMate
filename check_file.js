// ファイルの存在を確認するだけの単純なスクリプト

try {
    // ファイルシステムオブジェクトを作成
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    
    // ファイルパスを指定
    var filePath = "C:\\Users\\Orinr\\Documents\\sample_contract.docx";
    
    // ファイルが存在するか確認
    if (fso.FileExists(filePath)) {
        WScript.Echo("ファイルが存在します: " + filePath);
        
        // ファイルの情報を取得
        var file = fso.GetFile(filePath);
        WScript.Echo("ファイルサイズ: " + file.Size + " バイト");
        WScript.Echo("作成日時: " + file.DateCreated);
        WScript.Echo("最終更新日時: " + file.DateLastModified);
    } else {
        WScript.Echo("ファイルが存在しません: " + filePath);
    }
    
} catch (error) {
    WScript.Echo("エラーが発生しました: " + error);
}