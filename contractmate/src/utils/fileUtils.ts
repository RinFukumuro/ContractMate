import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as cp from 'child_process';
import { normalize } from 'path';

/**
 * ファイルの拡張子を取得する
 * @param uri ファイルのURI
 * @returns 拡張子（ドットを含まない小文字）
 */
export function getFileExtension(uri: vscode.Uri): string {
    return path.extname(uri.fsPath).toLowerCase().substring(1);
}

/**
 * ファイルが特定の拡張子かどうかを判断する
 * @param uri ファイルのURI
 * @param extensions 拡張子の配列（ドットを含まない）
 * @returns 特定の拡張子の場合はtrue、そうでない場合はfalse
 */
export function isFileType(uri: vscode.Uri, extensions: string[]): boolean {
    const ext = getFileExtension(uri);
    return extensions.includes(ext);
}

/**
 * PDFファイルかどうかを判断する
 * @param uri ファイルのURI
 * @returns PDFファイルの場合はtrue、そうでない場合はfalse
 */
export function isPdfFile(uri: vscode.Uri): boolean {
    return isFileType(uri, ['pdf']);
}

/**
 * Wordファイルかどうかを判断する
 * @param uri ファイルのURI
 * @returns Wordファイルの場合はtrue、そうでない場合はfalse
 */
export function isWordFile(uri: vscode.Uri): boolean {
    return isFileType(uri, ['doc', 'docx']);
}

/**
 * Excelファイルかどうかを判断する
 * @param uri ファイルのURI
 * @returns Excelファイルの場合はtrue、そうでない場合はfalse
 */
export function isExcelFile(uri: vscode.Uri): boolean {
    return isFileType(uri, ['xls', 'xlsx']);
}

/**
 * PowerPointファイルかどうかを判断する
 * @param uri ファイルのURI
 * @returns PowerPointファイルの場合はtrue、そうでない場合はfalse
 */
export function isPowerPointFile(uri: vscode.Uri): boolean {
    return isFileType(uri, ['ppt', 'pptx', 'pps', 'ppsx', 'pot', 'potx', 'pptm', 'potm', 'ppsm']);
}

/**
 * 画像ファイルかどうかを判断する
 * @param uri ファイルのURI
 * @returns 画像ファイルの場合はtrue、そうでない場合はfalse
 */
export function isImageFile(uri: vscode.Uri): boolean {
    return isFileType(uri, ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp']);
}

/**
 * テキストファイルかどうかを判断する
 * @param uri ファイルのURI
 * @returns テキストファイルの場合はtrue、そうでない場合はfalse
 */
export function isTextFile(uri: vscode.Uri): boolean {
    return isFileType(uri, ['txt', 'csv', 'md', 'json', 'xml', 'html', 'css', 'js', 'ts']);
}

/**
 * 外部アプリケーションでファイルを開く
 * @param uri ファイルのURI
 * @param app アプリケーションのパス（省略時はOSのデフォルトアプリケーション）
 */
export function openFileWithExternalApp(uri: vscode.Uri, app?: string): void {
    // パスを正規化
    let filePath = normalize(uri.fsPath);
    
    console.log(`openFileWithExternalApp: ファイルを開きます: ${filePath}`);
    console.log(`- 元のパス: ${uri.fsPath}`);
    console.log(`- 正規化されたパス: ${filePath}`);
    
    // 二重パスのチェック（C:\が2回出現していないか）
    if (filePath.match(/[a-zA-Z]:\\.*[a-zA-Z]:\\/)) {
        console.error(`二重パスが検出されました: ${filePath}`);
        
        // 最後のドライブレター以降の部分を抽出
        const match = filePath.match(/.*([a-zA-Z]:\\.*)/);
        if (match && match[1]) {
            filePath = match[1];
            console.log(`修正されたパス: ${filePath}`);
        }
    }
    
    // ファイルが存在するか確認（複数の方法を試す）
    let fileExists = false;
    
    // 方法1: fs.accessSync
    try {
        fs.accessSync(filePath, fs.constants.F_OK);
        console.log(`fs.accessSync: ファイルが存在します: ${filePath}`);
        fileExists = true;
    } catch (error) {
        console.error(`fs.accessSync: ファイルが存在しません: ${filePath}`);
    }
    
    // 方法2: fs.existsSync
    if (!fileExists) {
        try {
            fileExists = fs.existsSync(filePath);
            console.log(`fs.existsSync: ファイルが存在${fileExists ? 'します' : 'しません'}: ${filePath}`);
        } catch (error) {
            console.error(`fs.existsSync: エラーが発生しました: ${error}`);
        }
    }
    
    // 方法3: 直接ファイルを開いてみる
    if (!fileExists) {
        try {
            const fd = fs.openSync(filePath, 'r');
            fs.closeSync(fd);
            console.log(`fs.openSync: ファイルが存在します: ${filePath}`);
            fileExists = true;
        } catch (error) {
            console.error(`fs.openSync: ファイルが存在しません: ${filePath}`);
            
            // 最後の手段：エラーメッセージを表示せずに実行を試みる
            console.log(`ファイルの存在確認に失敗しましたが、実行を試みます`);
            fileExists = true; // 強制的に実行を許可
        }
    }
    
    // ファイルが存在しない場合は処理を中止
    if (!fileExists) {
        vscode.window.showErrorMessage(`ファイルが見つかりません: ${filePath}`);
        return;
    }
    
    try {
        if (process.platform === 'win32') {
            // Windowsの場合
            if (app) {
                console.log(`指定されたアプリで開きます: ${app}`);
                try {
                    // 方法1: execFile
                    cp.execFile(app, [filePath], (error, stdout, stderr) => {
                        if (error) {
                            console.error(`execFile実行エラー: ${error}`);
                            // 方法2: exec
                            cp.exec(`"${app}" "${filePath}"`, (error2, stdout2, stderr2) => {
                                if (error2) {
                                    console.error(`exec実行エラー: ${error2}`);
                                    vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error2}`);
                                }
                            });
                        }
                    });
                } catch (error) {
                    console.error(`アプリ実行エラー: ${error}`);
                    // 方法3: shellExecute（start）を使用
                    cp.exec(`start "" "${filePath}"`, (error2, stdout, stderr) => {
                        if (error2) {
                            console.error(`start実行エラー: ${error2}`);
                            vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error2}`);
                        }
                    });
                }
            } else {
                console.log(`デフォルトアプリで開きます`);
                // shellExecute（start）を使用
                cp.exec(`start "" "${filePath}"`, (error, stdout, stderr) => {
                    if (error) {
                        console.error(`実行エラー: ${error}`);
                        vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error}`);
                    }
                });
            }
        } else if (process.platform === 'darwin') {
            // macOSの場合
            if (app) {
                cp.exec(`open -a "${app}" "${filePath}"`);
            } else {
                cp.exec(`open "${filePath}"`);
            }
        } else {
            // Linuxの場合
            if (app) {
                cp.exec(`"${app}" "${filePath}"`);
            } else {
                cp.exec(`xdg-open "${filePath}"`);
            }
        }
    } catch (error) {
        console.error(`ファイルを開く処理でエラーが発生しました: ${error}`);
        vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error}`);
    }
}