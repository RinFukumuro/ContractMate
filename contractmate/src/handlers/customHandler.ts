import * as vscode from 'vscode';
import { openFileWithExternalApp } from '../utils/fileUtils';

/**
 * カスタム拡張子のファイルを処理するハンドラークラス
 */
export class CustomHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // カスタム拡張子のファイルを開くコマンドを登録
        const openCustomCommand = vscode.commands.registerCommand(
            'contractmate.openCustom',
            (uri: vscode.Uri) => {
                this.openCustom(uri);
            }
        );

        // カスタムアプリケーションを設定するコマンドを登録
        const configureCustomAppCommand = vscode.commands.registerCommand(
            'contractmate.configureCustomApp',
            () => {
                this.configureCustomApp();
            }
        );

        return [openCustomCommand, configureCustomAppCommand];
    }

    /**
     * カスタム拡張子のファイルを開く
     * @param uri ファイルのURI
     */
    private async openCustom(uri: vscode.Uri): Promise<void> {
        try {
            // ファイルの拡張子を取得
            const ext = uri.fsPath.split('.').pop()?.toLowerCase() || '';
            
            // 設定からカスタム拡張子の設定を取得
            const config = vscode.workspace.getConfiguration('contractmate');
            const customExtensions = config.get<{ [key: string]: string }>('customExtensions') || {};
            
            // 拡張子に対応するアプリケーションのパスを取得
            const appPath = customExtensions[ext];
            
            if (appPath) {
                // 外部アプリケーションでファイルを開く
                openFileWithExternalApp(uri, appPath);
                vscode.window.showInformationMessage(`${appPath}で開いています: ${uri.fsPath}`);
            } else {
                // 設定がない場合は、設定を促すメッセージを表示
                const configure = '設定する';
                const useDefault = 'デフォルトアプリで開く';
                
                const selected = await vscode.window.showInformationMessage(
                    `拡張子「${ext}」に対応するアプリケーションが設定されていません。`,
                    configure,
                    useDefault
                );
                
                if (selected === configure) {
                    // 拡張子に対応するアプリケーションを設定
                    this.configureCustomAppForExtension(ext);
                } else if (selected === useDefault) {
                    // デフォルトアプリケーションでファイルを開く
                    openFileWithExternalApp(uri);
                    vscode.window.showInformationMessage(`デフォルトアプリケーションで開いています: ${uri.fsPath}`);
                }
            }
        } catch (error) {
            vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error}`);
        }
    }

    /**
     * カスタムアプリケーションを設定する
     */
    private async configureCustomApp(): Promise<void> {
        // 拡張子を入力
        const ext = await vscode.window.showInputBox({
            prompt: '設定する拡張子を入力してください（例: pdf）',
            placeHolder: '拡張子（ドットなし）'
        });
        
        if (!ext) {
            return;
        }
        
        // 拡張子に対応するアプリケーションを設定
        this.configureCustomAppForExtension(ext);
    }

    /**
     * 特定の拡張子に対応するアプリケーションを設定する
     * @param ext 拡張子
     */
    private async configureCustomAppForExtension(ext: string): Promise<void> {
        // アプリケーションのパスを入力
        const appPath = await vscode.window.showInputBox({
            prompt: `拡張子「${ext}」に対応するアプリケーションのパスを入力してください`,
            placeHolder: 'アプリケーションのパス（例: C:\\Program Files\\App\\app.exe）'
        });
        
        if (!appPath) {
            return;
        }
        
        // 設定を更新
        const config = vscode.workspace.getConfiguration('contractmate');
        const customExtensions = config.get<{ [key: string]: string }>('customExtensions') || {};
        
        customExtensions[ext] = appPath;
        
        await config.update('customExtensions', customExtensions, vscode.ConfigurationTarget.Global);
        
        vscode.window.showInformationMessage(`拡張子「${ext}」に対応するアプリケーションを設定しました: ${appPath}`);
    }
}