import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import { openFileWithExternalApp } from '../utils/fileUtils';

/**
 * PowerPointファイルを処理するハンドラークラス
 */
export class PowerPointHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // PowerPointファイルを開くコマンドを登録
        const openPowerPointCommand = vscode.commands.registerCommand(
            'contractmate.openPowerPoint',
            (uri: vscode.Uri, slide?: number) => {
                this.openPowerPoint(uri, slide);
            }
        );

        // 特定のスライドでPowerPointファイルを開くコマンドを登録
        const openPowerPointAtSlideCommand = vscode.commands.registerCommand(
            'contractmate.openPowerPointAtSlide',
            (uri: vscode.Uri, slide: number) => {
                this.openPowerPoint(uri, slide);
            }
        );

        return [openPowerPointCommand, openPowerPointAtSlideCommand];
    }

    /**
     * PowerPointファイルを開く
     * @param uri PowerPointファイルのuri
     * @param slide スライド番号（省略可）
     */
    private openPowerPoint(uri: vscode.Uri, slide?: number): void {
        try {
            console.log(`openPowerPoint: PowerPointファイルを開きます: ${uri.fsPath}`);
            
            // URIの文字列表現を取得
            const uriString = uri.toString();
            
            // URIからスライド番号を抽出（指定されていない場合）
            if (slide === undefined) {
                // URIの文字列表現からスライド番号を抽出
                // #slide=N または #N 形式のチェック
                const slideMatch = uriString.match(/#slide[=%]?(\d+)/i) || uriString.match(/#page[=%]?(\d+)/i) || uriString.match(/#(\d+)$/);
                if (slideMatch && slideMatch[1]) {
                    slide = parseInt(slideMatch[1], 10);
                    console.log(`URIからスライド番号を抽出しました: ${slide}`);
                }
                
                // ?slide=N または ?page=N 形式のチェック（クエリパラメータ）
                if (slide === undefined) {
                    const queryMatch = uriString.match(/\?slide=(\d+)/i) || uriString.match(/\?page=(\d+)/i);
                    if (queryMatch && queryMatch[1]) {
                        slide = parseInt(queryMatch[1], 10);
                        console.log(`クエリパラメータからスライド番号を抽出しました: ${slide}`);
                    }
                }
            }
            
            // 設定からPowerPointアプリケーションのパスを取得
            const config = vscode.workspace.getConfiguration('contractmate');
            const powerPointAppPath = config.get<string>('powerPointAppPath');
            
            // スライド番号が指定されている場合
            if (slide !== undefined && slide > 0) {
                console.log(`PowerPointをスライド ${slide} で開きます: ${uri.fsPath}`);
                
                if (process.platform === 'win32') {
                    // Windowsの場合、コマンドラインオプションを使用してスライド番号を指定
                    if (powerPointAppPath) {
                        // PowerPointのパスが設定されている場合
                        console.log(`PowerPointのパスが設定されています: ${powerPointAppPath}`);
                        
                        // コマンドラインオプションを使用してPowerPointを開く
                        // /s はファイルを開くオプション、/n はスライド番号を指定するオプション
                        const command = `"${powerPointAppPath}" /s "${uri.fsPath}" /n ${slide}`;
                        console.log(`実行するコマンド: ${command}`);
                        
                        cp.exec(command, (error, stdout, stderr) => {
                            if (error) {
                                console.error(`PowerPointの実行中にエラーが発生しました: ${error}`);
                                vscode.window.showErrorMessage(`PowerPointファイルを開けませんでした: ${error}`);
                                
                                // エラーが発生した場合は、通常の方法でファイルを開く
                                openFileWithExternalApp(uri, powerPointAppPath);
                            }
                        });
                    } else {
                        // PowerPointのパスが設定されていない場合は、通常の方法でファイルを開く
                        console.log(`PowerPointのパスが設定されていないため、通常の方法でファイルを開きます`);
                        openFileWithExternalApp(uri);
                        
                        // スライド番号が指定されている場合は、ユーザーに通知
                        vscode.window.showInformationMessage(`PowerPointのパスが設定されていないため、スライド ${slide} に自動的にジャンプできません。手動でスライドに移動してください。`);
                    }
                } else {
                    // Windows以外の場合は、通常の方法でファイルを開く
                    console.log(`Windows以外のプラットフォームでは、スライド番号指定はサポートされていません`);
                    openFileWithExternalApp(uri, powerPointAppPath);
                    
                    // スライド番号が指定されている場合は、ユーザーに通知
                    vscode.window.showInformationMessage(`このプラットフォームでは、PowerPointのスライド番号指定はサポートされていません。手動でスライド ${slide} に移動してください。`);
                }
            } else {
                // スライド番号が指定されていない場合は、通常の方法でファイルを開く
                console.log(`通常の方法でPowerPointファイルを開きます: ${uri.fsPath}`);
                openFileWithExternalApp(uri, powerPointAppPath);
            }
            
            vscode.window.showInformationMessage(`PowerPointで開いています: ${uri.fsPath}${slide ? ` (スライド ${slide})` : ''}`);
        } catch (error) {
            vscode.window.showErrorMessage(`PowerPointファイルを開けませんでした: ${error}`);
        }
    }
}