import * as vscode from 'vscode';
import * as path from 'path';

/**
 * 画像ファイルを処理するハンドラークラス
 */
export class ImageHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // 画像ファイルを開くコマンドを登録
        const openImageCommand = vscode.commands.registerCommand(
            'contractmate.openImage',
            (uri: vscode.Uri) => {
                this.openImage(uri);
            }
        );

        // 画像ファイルのカスタムエディタプロバイダを登録
        const imageEditorProvider = vscode.window.registerCustomEditorProvider(
            'contractmate.imageViewer',
            new ImageEditorProvider(this.context),
            {
                webviewOptions: {
                    retainContextWhenHidden: true,
                },
                supportsMultipleEditorsPerDocument: false,
            }
        );

        return [openImageCommand, imageEditorProvider];
    }

    /**
     * 画像ファイルを開く
     * @param uri 画像ファイルのURI
     */
    private async openImage(uri: vscode.Uri): Promise<void> {
        try {
            // カスタムエディタで画像を開く
            await vscode.commands.executeCommand('vscode.openWith', uri, 'contractmate.imageViewer');
        } catch (error) {
            vscode.window.showErrorMessage(`画像ファイルを開けませんでした: ${error}`);
        }
    }
}

/**
 * 画像ファイルのカスタムエディタプロバイダ
 */
class ImageEditorProvider implements vscode.CustomReadonlyEditorProvider {
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * カスタムエディタを開く
     */
    async openCustomDocument(
        uri: vscode.Uri,
        _openContext: vscode.CustomDocumentOpenContext,
        _token: vscode.CancellationToken
    ): Promise<vscode.CustomDocument> {
        return { uri, dispose: () => {} };
    }

    /**
     * WebViewパネルを解決する
     */
    async resolveCustomEditor(
        document: vscode.CustomDocument,
        webviewPanel: vscode.WebviewPanel,
        _token: vscode.CancellationToken
    ): Promise<void> {
        // WebViewの設定
        webviewPanel.webview.options = {
            enableScripts: true,
            localResourceRoots: [
                vscode.Uri.file(path.dirname(document.uri.fsPath)),
                this.context.extensionUri
            ]
        };

        // 画像データを取得
        const imageData = await this.getImageData(document.uri);
        
        // WebViewのHTMLを設定
        webviewPanel.webview.html = this.getHtmlForWebview(webviewPanel.webview, document.uri, imageData);
    }

    /**
     * 画像ファイルのデータを取得する
     */
    private async getImageData(uri: vscode.Uri): Promise<Uint8Array> {
        return await vscode.workspace.fs.readFile(uri);
    }

    /**
     * WebViewのHTMLを生成する
     */
    private getHtmlForWebview(webview: vscode.Webview, uri: vscode.Uri, imageData: Uint8Array): string {
        // 画像のMIMEタイプを取得
        const mimeType = this.getMimeType(uri.fsPath);
        
        // 画像データをBase64エンコード
        const imageDataBase64 = Buffer.from(imageData).toString('base64');
        
        // ファイル名を取得
        const fileName = path.basename(uri.fsPath);

        return `
            <!DOCTYPE html>
            <html lang="ja">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>${fileName}</title>
                <style>
                    body {
                        margin: 0;
                        padding: 20px;
                        background-color: var(--vscode-editor-background);
                        color: var(--vscode-editor-foreground);
                        font-family: var(--vscode-font-family);
                        font-size: var(--vscode-font-size);
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;
                        height: 100vh;
                        overflow: auto;
                    }
                    .image-container {
                        max-width: 100%;
                        max-height: 90vh;
                        overflow: auto;
                        text-align: center;
                    }
                    .image-container img {
                        max-width: 100%;
                        max-height: 100%;
                        object-fit: contain;
                    }
                    .file-info {
                        margin-top: 10px;
                        font-size: 12px;
                        color: var(--vscode-descriptionForeground);
                    }
                    .controls {
                        margin-top: 10px;
                        display: flex;
                        gap: 10px;
                    }
                    button {
                        background-color: var(--vscode-button-background);
                        color: var(--vscode-button-foreground);
                        border: none;
                        padding: 5px 10px;
                        cursor: pointer;
                        border-radius: 2px;
                    }
                    button:hover {
                        background-color: var(--vscode-button-hoverBackground);
                    }
                </style>
            </head>
            <body>
                <div class="image-container">
                    <img src="data:${mimeType};base64,${imageDataBase64}" alt="${fileName}" id="image">
                </div>
                <div class="file-info">
                    ${fileName}
                </div>
                <div class="controls">
                    <button id="zoomIn">拡大</button>
                    <button id="zoomOut">縮小</button>
                    <button id="resetZoom">リセット</button>
                </div>

                <script>
                    // 画像の拡大・縮小機能
                    const image = document.getElementById('image');
                    const zoomInButton = document.getElementById('zoomIn');
                    const zoomOutButton = document.getElementById('zoomOut');
                    const resetZoomButton = document.getElementById('resetZoom');
                    
                    let scale = 1;
                    
                    zoomInButton.addEventListener('click', () => {
                        scale *= 1.2;
                        updateScale();
                    });
                    
                    zoomOutButton.addEventListener('click', () => {
                        scale /= 1.2;
                        updateScale();
                    });
                    
                    resetZoomButton.addEventListener('click', () => {
                        scale = 1;
                        updateScale();
                    });
                    
                    function updateScale() {
                        image.style.transform = \`scale(\${scale})\`;
                    }
                </script>
            </body>
            </html>
        `;
    }

    /**
     * ファイルの拡張子からMIMEタイプを取得する
     */
    private getMimeType(filePath: string): string {
        const ext = path.extname(filePath).toLowerCase();
        
        switch (ext) {
            case '.png':
                return 'image/png';
            case '.jpg':
            case '.jpeg':
                return 'image/jpeg';
            case '.gif':
                return 'image/gif';
            case '.bmp':
                return 'image/bmp';
            case '.webp':
                return 'image/webp';
            default:
                return 'image/png';
        }
    }
}