import * as vscode from 'vscode';
import * as path from 'path';

/**
 * テキストファイルを処理するハンドラークラス
 */
export class TextHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // テキストファイルを開くコマンドを登録
        const openTextCommand = vscode.commands.registerCommand(
            'contractmate.openText',
            (uri: vscode.Uri, line?: number) => {
                this.openText(uri, line);
            }
        );

        // 特定の行でテキストファイルを開くコマンドを登録
        const openTextAtLineCommand = vscode.commands.registerCommand(
            'contractmate.openTextAtLine',
            (uri: vscode.Uri, line: number) => {
                this.openText(uri, line);
            }
        );

        return [openTextCommand, openTextAtLineCommand];
    }

    /**
     * テキストファイルを開く
     * @param uri テキストファイルのURI
     * @param line 開始行番号（省略可）
     */
    private async openText(uri: vscode.Uri, line?: number): Promise<void> {
        try {
            console.log(`openText: テキストファイルを開きます: ${uri.fsPath}`);
            
            // URIの文字列表現を取得
            const uriString = uri.toString();
            
            // URIからページ番号（行番号）を抽出（指定されていない場合）
            if (line === undefined) {
                // URIの文字列表現から行番号を抽出
                // #line=N または #N 形式のチェック
                const lineMatch = uriString.match(/#line[=%]?(\d+)/i) || uriString.match(/#(\d+)$/);
                if (lineMatch && lineMatch[1]) {
                    line = parseInt(lineMatch[1], 10);
                    console.log(`URIから行番号を抽出しました: ${line}`);
                }
                
                // ?line=N 形式のチェック（クエリパラメータ）
                if (line === undefined && uriString.includes('?line=')) {
                    const queryMatch = uriString.match(/\?line=(\d+)/i);
                    if (queryMatch && queryMatch[1]) {
                        line = parseInt(queryMatch[1], 10);
                        console.log(`クエリパラメータから行番号を抽出しました: ${line}`);
                    }
                }
            }
            
            // VSCodeのデフォルトエディタでファイルを開く
            const document = await vscode.workspace.openTextDocument(uri);
            const editor = await vscode.window.showTextDocument(document);
            
            // 行番号が指定されている場合、その行にジャンプ
            if (line !== undefined && line > 0) {
                // 行番号は0から始まるため、1を引く
                const lineIndex = Math.min(line - 1, document.lineCount - 1);
                
                // 指定された行の範囲を取得
                const lineRange = document.lineAt(lineIndex).range;
                
                // エディタの選択範囲を設定
                editor.selection = new vscode.Selection(lineRange.start, lineRange.start);
                
                // 指定された行が表示されるようにスクロール
                editor.revealRange(lineRange, vscode.TextEditorRevealType.InCenter);
                
                console.log(`行 ${line} にジャンプしました`);
            }
        } catch (error) {
            vscode.window.showErrorMessage(`テキストファイルを開けませんでした: ${error}`);
        }
    }
}