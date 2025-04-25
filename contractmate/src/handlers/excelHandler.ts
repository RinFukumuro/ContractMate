import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import { openFileWithExternalApp } from '../utils/fileUtils';

/**
 * Excelファイルを処理するハンドラークラス
 */
export class ExcelHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // Excelファイルを開くコマンドを登録
        const openExcelCommand = vscode.commands.registerCommand(
            'contractmate.openExcel',
            (uri: vscode.Uri, sheet?: string, cell?: string) => {
                this.openExcel(uri, sheet, cell);
            }
        );

        // 特定のシートとセルでExcelファイルを開くコマンドを登録
        const openExcelAtSheetCellCommand = vscode.commands.registerCommand(
            'contractmate.openExcelAtSheetCell',
            (uri: vscode.Uri, sheet: string, cell?: string) => {
                this.openExcel(uri, sheet, cell);
            }
        );

        return [openExcelCommand, openExcelAtSheetCellCommand];
    }

    /**
     * Excelファイルを開く
     * @param uri Excelファイルのuri
     * @param sheet シート名（省略可）
     * @param cell セル位置（省略可）
     */
    private openExcel(uri: vscode.Uri, sheet?: string, cell?: string): void {
        try {
            console.log(`openExcel: Excelファイルを開きます: ${uri.fsPath}`);
            
            // URIの文字列表現を取得
            const uriString = uri.toString();
            
            // URIからシート名とセル位置を抽出（指定されていない場合）
            if (!sheet && !cell) {
                // #sheet=SheetName!CellRef 形式のチェック
                const sheetCellMatch = uriString.match(/#sheet=([^!]+)!([A-Z0-9]+)/i);
                if (sheetCellMatch && sheetCellMatch[1] && sheetCellMatch[2]) {
                    sheet = sheetCellMatch[1];
                    cell = sheetCellMatch[2];
                    console.log(`URIからシート名とセル位置を抽出しました: シート=${sheet}, セル=${cell}`);
                }
                // #sheet=SheetName 形式のチェック
                else {
                    const sheetMatch = uriString.match(/#sheet=([^!&]+)/i);
                    if (sheetMatch && sheetMatch[1]) {
                        sheet = sheetMatch[1];
                        console.log(`URIからシート名を抽出しました: シート=${sheet}`);
                    }
                }
                
                // #cell=CellRef 形式のチェック（シート名が既に設定されている場合）
                if (sheet && !cell) {
                    const cellMatch = uriString.match(/#cell=([A-Z0-9]+)/i);
                    if (cellMatch && cellMatch[1]) {
                        cell = cellMatch[1];
                        console.log(`URIからセル位置を抽出しました: セル=${cell}`);
                    }
                }
                
                // ?sheet=SheetName&cell=CellRef 形式のチェック（クエリパラメータ）
                if (!sheet && !cell) {
                    const querySheetMatch = uriString.match(/\?sheet=([^!&]+)/i);
                    const queryCellMatch = uriString.match(/[?&]cell=([A-Z0-9]+)/i);
                    
                    if (querySheetMatch && querySheetMatch[1]) {
                        sheet = querySheetMatch[1];
                        console.log(`クエリパラメータからシート名を抽出しました: シート=${sheet}`);
                    }
                    
                    if (queryCellMatch && queryCellMatch[1]) {
                        cell = queryCellMatch[1];
                        console.log(`クエリパラメータからセル位置を抽出しました: セル=${cell}`);
                    }
                }
                
                // #page=N 形式のチェック（ページ番号をシート番号として扱う）
                if (!sheet) {
                    const pageMatch = uriString.match(/#page=(\d+)/i) || uriString.match(/\?page=(\d+)/i);
                    if (pageMatch && pageMatch[1]) {
                        const sheetNumber = parseInt(pageMatch[1], 10);
                        // シート番号を1から始まるインデックスとして扱う
                        if (sheetNumber > 0) {
                            // シート番号をシート名として使用（Sheet1, Sheet2, ...）
                            sheet = `Sheet${sheetNumber}`;
                            console.log(`ページ番号からシート名を生成しました: シート=${sheet}`);
                        }
                    }
                }
            }
            
            // 設定からExcelアプリケーションのパスを取得
            const config = vscode.workspace.getConfiguration('contractmate');
            const excelAppPath = config.get<string>('excelAppPath');
            
            // シート名またはセル位置が指定されている場合
            if (sheet || cell) {
                console.log(`Excelをシート=${sheet || 'デフォルト'}, セル=${cell || 'デフォルト'}で開きます: ${uri.fsPath}`);
                
                if (process.platform === 'win32') {
                    // Windowsの場合、コマンドラインオプションを使用してシート名とセル位置を指定
                    if (excelAppPath) {
                        // Excelのパスが設定されている場合
                        console.log(`Excelのパスが設定されています: ${excelAppPath}`);
                        
                        // シート名とセル位置を組み合わせる
                        let sheetCell = '';
                        if (sheet) {
                            sheetCell = sheet;
                            if (cell) {
                                sheetCell += `!${cell}`;
                            }
                        }
                        
                        // コマンドラインオプションを使用してExcelを開く
                        // /e はファイルを開くオプション、/o はシート名とセル位置を指定するオプション
                        let command = `"${excelAppPath}" /e "${uri.fsPath}"`;
                        if (sheetCell) {
                            command += ` /o "${sheetCell}"`;
                        }
                        console.log(`実行するコマンド: ${command}`);
                        
                        cp.exec(command, (error, stdout, stderr) => {
                            if (error) {
                                console.error(`Excelの実行中にエラーが発生しました: ${error}`);
                                vscode.window.showErrorMessage(`Excelファイルを開けませんでした: ${error}`);
                                
                                // エラーが発生した場合は、通常の方法でファイルを開く
                                openFileWithExternalApp(uri, excelAppPath);
                            }
                        });
                    } else {
                        // Excelのパスが設定されていない場合は、通常の方法でファイルを開く
                        console.log(`Excelのパスが設定されていないため、通常の方法でファイルを開きます`);
                        openFileWithExternalApp(uri);
                        
                        // シート名またはセル位置が指定されている場合は、ユーザーに通知
                        let message = 'Excelのパスが設定されていないため、';
                        if (sheet && cell) {
                            message += `シート "${sheet}" のセル "${cell}" に自動的にジャンプできません。`;
                        } else if (sheet) {
                            message += `シート "${sheet}" に自動的にジャンプできません。`;
                        } else if (cell) {
                            message += `セル "${cell}" に自動的にジャンプできません。`;
                        }
                        message += '手動で移動してください。';
                        vscode.window.showInformationMessage(message);
                    }
                } else {
                    // Windows以外の場合は、通常の方法でファイルを開く
                    console.log(`Windows以外のプラットフォームでは、シート名とセル位置の指定はサポートされていません`);
                    openFileWithExternalApp(uri, excelAppPath);
                    
                    // シート名またはセル位置が指定されている場合は、ユーザーに通知
                    let message = 'このプラットフォームでは、Excelのシート名とセル位置の指定はサポートされていません。手動で';
                    if (sheet && cell) {
                        message += `シート "${sheet}" のセル "${cell}" に移動してください。`;
                    } else if (sheet) {
                        message += `シート "${sheet}" に移動してください。`;
                    } else if (cell) {
                        message += `セル "${cell}" に移動してください。`;
                    }
                    vscode.window.showInformationMessage(message);
                }
            } else {
                // シート名もセル位置も指定されていない場合は、通常の方法でファイルを開く
                console.log(`通常の方法でExcelファイルを開きます: ${uri.fsPath}`);
                openFileWithExternalApp(uri, excelAppPath);
            }
            
            // 情報メッセージを表示
            let infoMessage = `Excelで開いています: ${uri.fsPath}`;
            if (sheet && cell) {
                infoMessage += ` (シート "${sheet}" のセル "${cell}")`;
            } else if (sheet) {
                infoMessage += ` (シート "${sheet}")`;
            } else if (cell) {
                infoMessage += ` (セル "${cell}")`;
            }
            vscode.window.showInformationMessage(infoMessage);
        } catch (error) {
            vscode.window.showErrorMessage(`Excelファイルを開けませんでした: ${error}`);
        }
    }
}