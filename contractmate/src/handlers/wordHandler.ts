import * as vscode from 'vscode';
import * as path from 'path';
import * as cp from 'child_process';
import { openFileWithExternalApp } from '../utils/fileUtils';

/**
 * Wordファイルを処理するハンドラークラス
 */
export class WordHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // Wordファイルを開くコマンドを登録
        const openWordCommand = vscode.commands.registerCommand(
            'contractmate.openWord',
            (uri: vscode.Uri, page?: number, bookmark?: string) => {
                this.openWord(uri, page, bookmark);
            }
        );

        // 特定のページでWordファイルを開くコマンドを登録
        const openWordAtPageCommand = vscode.commands.registerCommand(
            'contractmate.openWordAtPage',
            (uri: vscode.Uri, page: number) => {
                this.openWord(uri, page);
            }
        );

        // 特定のブックマークでWordファイルを開くコマンドを登録
        const openWordAtBookmarkCommand = vscode.commands.registerCommand(
            'contractmate.openWordAtBookmark',
            (uri: vscode.Uri, bookmark: string) => {
                this.openWord(uri, undefined, bookmark);
            }
        );

        return [openWordCommand, openWordAtPageCommand, openWordAtBookmarkCommand];
    }

    /**
     * Wordファイルを開く
     * @param uri Wordファイルのuri
     * @param page ページ番号（省略可）
     * @param bookmark ブックマーク名（省略可）
     */
    private openWord(uri: vscode.Uri, page?: number, bookmark?: string): void {
        try {
            console.log(`openWord: Wordファイルを開きます: ${uri.fsPath}`);
            
            // URIの文字列表現を取得
            const uriString = uri.toString();
            
            // URIからページ番号またはブックマーク名を抽出（指定されていない場合）
            if (page === undefined && bookmark === undefined) {
                // #page=N 形式のチェック
                const pageMatch = uriString.match(/#page[=%]?(\d+)/i) || uriString.match(/#(\d+)$/);
                if (pageMatch && pageMatch[1]) {
                    page = parseInt(pageMatch[1], 10);
                    console.log(`URIからページ番号を抽出しました: ${page}`);
                }
                
                // ?page=N 形式のチェック（クエリパラメータ）
                if (page === undefined && uriString.includes('?page=')) {
                    const queryMatch = uriString.match(/\?page=(\d+)/i);
                    if (queryMatch && queryMatch[1]) {
                        page = parseInt(queryMatch[1], 10);
                        console.log(`クエリパラメータからページ番号を抽出しました: ${page}`);
                    }
                }
                
                // #bookmark=Name 形式のチェック
                if (bookmark === undefined) {
                    const bookmarkMatch = uriString.match(/#bookmark=([^&]+)/i);
                    if (bookmarkMatch && bookmarkMatch[1]) {
                        bookmark = bookmarkMatch[1];
                        console.log(`URIからブックマーク名を抽出しました: ${bookmark}`);
                    }
                }
                
                // ?bookmark=Name 形式のチェック（クエリパラメータ）
                if (bookmark === undefined && uriString.includes('?bookmark=')) {
                    const queryMatch = uriString.match(/\?bookmark=([^&]+)/i);
                    if (queryMatch && queryMatch[1]) {
                        bookmark = queryMatch[1];
                        console.log(`クエリパラメータからブックマーク名を抽出しました: ${bookmark}`);
                    }
                }
            }
            
            // 設定からWordアプリケーションのパスを取得
            const config = vscode.workspace.getConfiguration('contractmate');
            const wordAppPath = config.get<string>('wordAppPath');
            
            // ページ番号またはブックマーク名が指定されている場合
            if ((page !== undefined && page > 0) || (bookmark !== undefined && bookmark.length > 0)) {
                if (page !== undefined) {
                    console.log(`Wordをページ ${page} で開きます: ${uri.fsPath}`);
                }
                if (bookmark !== undefined) {
                    console.log(`Wordをブックマーク "${bookmark}" で開きます: ${uri.fsPath}`);
                }
                
                if (process.platform === 'win32') {
                    // Windowsの場合、コマンドラインオプションを使用してページ番号またはブックマーク名を指定
                    if (wordAppPath) {
                        // Wordのパスが設定されている場合
                        console.log(`Wordのパスが設定されています: ${wordAppPath}`);
                        
                        // コマンドラインオプションを使用してWordを開く
                        // /n は新しいウィンドウで開くオプション
                        let command = `"${wordAppPath}" /n "${uri.fsPath}"`;
                        console.log(`実行するコマンド: ${command}`);
                        
                        cp.exec(command, (error, stdout, stderr) => {
                            if (error) {
                                console.error(`Wordの実行中にエラーが発生しました: ${error}`);
                                vscode.window.showErrorMessage(`Wordファイルを開けませんでした: ${error}`);
                                
                                // エラーが発生した場合は、通常の方法でファイルを開く
                                openFileWithExternalApp(uri, wordAppPath);
                            } else {
                                // Wordが開かれた後、ページ番号またはブックマークにジャンプするためのVBAマクロを実行する
                                // 注意: これは実際には機能しない可能性があります。Wordのコマンドラインオプションには
                                // 直接ページ番号やブックマークを指定するオプションがないため。
                                if (page !== undefined || bookmark !== undefined) {
                                    vscode.window.showInformationMessage(
                                        `Wordファイルが開かれました。${page !== undefined ? `ページ ${page}` : `ブックマーク "${bookmark}"`} に手動で移動してください。`
                                    );
                                }
                            }
                        });
                    } else {
                        // Wordのパスが設定されていない場合は、通常の方法でファイルを開く
                        console.log(`Wordのパスが設定されていないため、通常の方法でファイルを開きます`);
                        openFileWithExternalApp(uri);
                        
                        // ページ番号またはブックマーク名が指定されている場合は、ユーザーに通知
                        let message = 'Wordのパスが設定されていないため、';
                        if (page !== undefined) {
                            message += `ページ ${page} に自動的にジャンプできません。`;
                        } else if (bookmark !== undefined) {
                            message += `ブックマーク "${bookmark}" に自動的にジャンプできません。`;
                        }
                        message += '手動で移動してください。';
                        vscode.window.showInformationMessage(message);
                    }
                } else {
                    // Windows以外の場合は、通常の方法でファイルを開く
                    console.log(`Windows以外のプラットフォームでは、ページ番号やブックマーク名の指定はサポートされていません`);
                    openFileWithExternalApp(uri, wordAppPath);
                    
                    // ページ番号またはブックマーク名が指定されている場合は、ユーザーに通知
                    let message = 'このプラットフォームでは、Wordのページ番号やブックマーク名の指定はサポートされていません。手動で';
                    if (page !== undefined) {
                        message += `ページ ${page} に移動してください。`;
                    } else if (bookmark !== undefined) {
                        message += `ブックマーク "${bookmark}" に移動してください。`;
                    }
                    vscode.window.showInformationMessage(message);
                }
            } else {
                // ページ番号もブックマーク名も指定されていない場合は、通常の方法でファイルを開く
                console.log(`通常の方法でWordファイルを開きます: ${uri.fsPath}`);
                openFileWithExternalApp(uri, wordAppPath);
            }
            
            // 情報メッセージを表示
            let infoMessage = `Wordで開いています: ${uri.fsPath}`;
            if (page !== undefined) {
                infoMessage += ` (ページ ${page})`;
            } else if (bookmark !== undefined) {
                infoMessage += ` (ブックマーク "${bookmark}")`;
            }
            vscode.window.showInformationMessage(infoMessage);
        } catch (error) {
            vscode.window.showErrorMessage(`Wordファイルを開けませんでした: ${error}`);
        }
    }
}