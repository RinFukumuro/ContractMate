import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';

/**
 * PDFファイルを処理するハンドラークラス
 */
export class PdfHandler {
    /**
     * コンストラクタ
     * @param context 拡張機能のコンテキスト
     */
    constructor(private readonly context: vscode.ExtensionContext) {}

    /**
     * ハンドラーを登録する
     */
    register(): vscode.Disposable[] {
        // PDFファイルを開くコマンドを登録
        const openPdfCommand = vscode.commands.registerCommand(
            'contractmate.openPdf',
            (uri: vscode.Uri, page?: number) => {
                this.openPdf(uri, page);
            }
        );

        // 特定のページでPDFファイルを開くコマンドを登録
        const openPdfAtPageCommand = vscode.commands.registerCommand(
            'contractmate.openPdfAtPage',
            (uri: vscode.Uri, page: number, originalUrl?: string) => {
                this.openPdf(uri, page, originalUrl);
            }
        );

        // PDFファイルのカスタムエディタプロバイダを登録
        const pdfEditorProvider = vscode.window.registerCustomEditorProvider(
            'contractmate.pdfViewer',
            new PdfEditorProvider(this.context),
            {
                webviewOptions: {
                    retainContextWhenHidden: true,
                },
                supportsMultipleEditorsPerDocument: false,
            }
        );

        return [openPdfCommand, openPdfAtPageCommand, pdfEditorProvider];
    }

    /**
     * PDFファイルを開く
     * @param uri PDFファイルのURI
     * @param page 開始ページ番号（省略可）
     */
    private async openPdf(uri: vscode.Uri, page?: number, originalUrl?: string): Promise<void> {
        // URIの文字列表現を取得
        const uriString = uri.toString();
        console.log(`openPdf: 元のURI: ${uriString}`);
        
        // 直接入力されたURLからページ番号を抽出（VSCodeのコマンドパレットから直接入力された場合）
        if (!originalUrl && uri.path) {
            // URIのパス部分からファイル名を抽出
            const fileName = path.basename(uri.path);
            
            // ファイル名にハッシュフラグメントが含まれているか確認
            if (fileName.includes('#')) {
                console.log(`ファイル名にハッシュフラグメントが含まれています: ${fileName}`);
                
                // ハッシュフラグメントを含むファイル名からオリジナルURLを再構築
                originalUrl = `file:///${uri.path}`;
                console.log(`再構築されたオリジナルURL: ${originalUrl}`);
            }
        }

        // デバッグ用：ページ番号が指定されている場合はログ出力
        if (page !== undefined) {
            console.log(`openPdf: 指定されたページ番号: ${page}`);
        }
        
        // URIからページ番号を抽出（指定されていない場合）
        if (page === undefined) {
            // URIの文字列表現からページ番号を抽出
            // #page=N または #page%3DN 形式のチェック
            const pageMatch = uriString.match(/#page[=%]?(\d+)/i) || uriString.match(/#page%3D(\d+)/i);
            if (pageMatch && pageMatch[1]) {
                page = parseInt(pageMatch[1], 10);
                console.log(`URIからページ番号を抽出しました: ${page}`);
                
                // ページ番号の検証は後で行うので、ここでは単に抽出するだけ
            }
        }
        try {
            // URIの詳細情報をログ出力
            console.log(`openPdf: URI詳細情報:`, {
                original: uri.toString(),
                scheme: uri.scheme,
                authority: uri.authority,
                path: uri.path,
                query: uri.query,
                fragment: uri.fragment,
                fsPath: uri.fsPath,
                originalUrl: originalUrl || 'なし'
            });
            
            // URIからページ番号を抽出（#page=3 または #999 形式）
            if (page === undefined) {
                // 1. 元のURL文字列から抽出（最も信頼性が高い）
                if (originalUrl) {
                    console.log(`元のURL文字列を解析: ${originalUrl}`);
                    
                    // デコード前のURLを出力（デバッグ用）
                    console.log(`デコード前のURL: ${originalUrl}`);
                    
                    // URLをデコード
                    let decodedUrl;
                    try {
                        decodedUrl = decodeURIComponent(originalUrl);
                        console.log(`デコード後のURL: ${decodedUrl}`);
                    } catch (e) {
                        console.log(`URLのデコードに失敗: ${e}`);
                        decodedUrl = originalUrl;
                    }
                    
                    // 1. #page=N 形式のチェック（デコード後のハッシュフラグメント）
                    const hashPageMatch = decodedUrl.match(/#page=(\d+)/i);
                    if (hashPageMatch && hashPageMatch[1]) {
                        page = parseInt(hashPageMatch[1], 10);
                        console.log(`デコードされたURLから抽出したページ番号(#page=N形式): ${page}`);
                    }
                    // 2. #page%3DN 形式のチェック（デコード前のURLエンコードされたハッシュフラグメント）
                    else {
                        const encodedHashPageMatch = originalUrl.match(/#page%3D(\d+)/i);
                        if (encodedHashPageMatch && encodedHashPageMatch[1]) {
                            page = parseInt(encodedHashPageMatch[1], 10);
                            console.log(`元のURLから抽出したページ番号(#page%3DN形式): ${page}`);
                        }
                        // 3. #N 形式のチェック（デコード後のハッシュフラグメント）
                        else {
                            const hashMatch = decodedUrl.match(/#(\d+)$/);
                            if (hashMatch && hashMatch[1]) {
                                page = parseInt(hashMatch[1], 10);
                                console.log(`デコードされたURLから抽出したページ番号(#N形式): ${page}`);
                            }
                            // 4. ?page=N 形式のチェック（デコード後のクエリパラメータ）
                            else {
                                const queryPageMatch = decodedUrl.match(/\?page=(\d+)/i);
                                if (queryPageMatch && queryPageMatch[1]) {
                                    page = parseInt(queryPageMatch[1], 10);
                                    console.log(`デコードされたURLから抽出したページ番号(?page=N形式): ${page}`);
                                }
                                // 5. ?page%3DN 形式のチェック（デコード前のURLエンコードされたクエリパラメータ）
                                else {
                                    const encodedQueryPageMatch = originalUrl.match(/\?page%3D(\d+)/i);
                                    if (encodedQueryPageMatch && encodedQueryPageMatch[1]) {
                                        page = parseInt(encodedQueryPageMatch[1], 10);
                                        console.log(`元のURLから抽出したページ番号(?page%3DN形式): ${page}`);
                                    }
                                }
                            }
                        }
                    }
                    
                    // ページ番号が見つかった場合、デバッグ情報を出力
                    if (page !== undefined) {
                        console.log(`ページ番号を抽出しました: ${page}`);
                    } else {
                        console.log(`ページ番号を抽出できませんでした`);
                    }
                }
                
                // 2. URI.fragmentを確認（VSCodeのURIオブジェクトの仕様により、fragmentプロパティが空の場合がある）
                if (page === undefined && uri.fragment) {
                    console.log(`URI.fragmentから直接抽出: ${uri.fragment}`);
                    
                    // page=N 形式のチェック
                    const pageMatch = uri.fragment.match(/^page=(\d+)$/i);
                    if (pageMatch && pageMatch[1]) {
                        page = parseInt(pageMatch[1], 10);
                        console.log(`URI.fragmentから抽出したページ番号(page=N形式): ${page}`);
                    }
                    // N 形式のチェック
                    else {
                        const numMatch = uri.fragment.match(/^(\d+)$/);
                        if (numMatch && numMatch[1]) {
                            page = parseInt(numMatch[1], 10);
                            console.log(`URI.fragmentから抽出したページ番号(N形式): ${page}`);
                        }
                    }
                }
                
                // 3. URI文字列全体から検索
                if (page === undefined) {
                    console.log(`URI文字列全体を解析: ${uriString}`);
                    
                    // #page=N 形式のチェック
                    const pageMatch = uriString.match(/#page=(\d+)/i);
                    if (pageMatch && pageMatch[1]) {
                        page = parseInt(pageMatch[1], 10);
                        console.log(`URI文字列から抽出したページ番号(#page=N形式): ${page}`);
                    }
                    // #page%3DN 形式のチェック（URLエンコードされたハッシュフラグメント）
                    else {
                        const encodedPageMatch = uriString.match(/#page%3D(\d+)/i);
                        if (encodedPageMatch && encodedPageMatch[1]) {
                            page = parseInt(encodedPageMatch[1], 10);
                            console.log(`URI文字列から抽出したページ番号(#page%3DN形式): ${page}`);
                        }
                        // #N 形式のチェック
                        else {
                            const hashMatch = uriString.match(/#(\d+)$/);
                            if (hashMatch && hashMatch[1]) {
                                page = parseInt(hashMatch[1], 10);
                                console.log(`URI文字列から抽出したページ番号(#N形式): ${page}`);
                            }
                        }
                    }
                }
            }
            
            // カスタムエディタでPDFを開く
            await vscode.commands.executeCommand('vscode.openWith', uri, 'contractmate.pdfViewer');
            
            // ページ番号が指定されている場合、そのページを表示するメッセージを送信
            if (page !== undefined) {
                // ページ番号が0以下の場合は1に設定（最初のページ）
                if (page <= 0) {
                    console.log(`指定されたページ番号 ${page} が無効です。最初のページを表示します。`);
                    page = 1;
                }
                // URIからハッシュ部分とクエリパラメータを削除（ページ情報を含まないクリーンなURI）
                const cleanUri = uri.toString().replace(/[?#].*$/, '');
                
                // 元のURI（変更なし）
                const originalUri = uri.toString();
                
                // ファイル名を取得
                const fileName = path.basename(uri.fsPath);
                
                // 現在のタイムスタンプを取得
                const timestamp = Date.now();
                
                // 複数のバリエーションでページジャンプ情報を保存
                console.log(`openPdf: ページ ${page} へのジャンプ情報を保存します`);
                
                // 1. クリーンURI（ハッシュとクエリパラメータを除去）
                PdfEditorProvider.pendingPageJumps.set(cleanUri, {
                    page: page,
                    timestamp: timestamp
                });
                console.log(`- クリーンURI: ${cleanUri}`);
                
                // 2. 元のURI（変更なし）
                PdfEditorProvider.pendingPageJumps.set(originalUri, {
                    page: page,
                    timestamp: timestamp
                });
                console.log(`- 元のURI: ${originalUri}`);
                
                // 3. ファイル名
                PdfEditorProvider.pendingPageJumpsByFileName.set(fileName, {
                    page: page,
                    timestamp: timestamp
                });
                console.log(`- ファイル名: ${fileName}`);
                
                // 4. デコードされたURI
                try {
                    const decodedUri = decodeURIComponent(cleanUri);
                    if (decodedUri !== cleanUri) {
                        PdfEditorProvider.pendingPageJumps.set(decodedUri, {
                            page: page,
                            timestamp: timestamp
                        });
                        console.log(`- デコードされたURI: ${decodedUri}`);
                    }
                } catch (e) {
                    console.log(`URIのデコードに失敗: ${e}`);
                }

                // 5. ファイルパスのみ（スキーム部分を除去）
                const filePathOnly = uri.fsPath;
                PdfEditorProvider.pendingPageJumps.set(filePathOnly, {
                    page: page,
                    timestamp: timestamp
                });
                console.log(`- ファイルパスのみ: ${filePathOnly}`);

                // 6. ファイルURIスキーム付き（file:///形式）
                const fileUri = `file:///${uri.fsPath.replace(/\\/g, '/')}`;
                PdfEditorProvider.pendingPageJumps.set(fileUri, {
                    page: page,
                    timestamp: timestamp
                });
                console.log(`- ファイルURIスキーム付き: ${fileUri}`);
                
                // 古いエントリを削除（5分以上経過したエントリ）
                const expirationTime = timestamp - 5 * 60 * 1000; // 5分前
                
                // URIマップの古いエントリを削除
                for (const [uri, info] of PdfEditorProvider.pendingPageJumps.entries()) {
                    if (info.timestamp < expirationTime) {
                        PdfEditorProvider.pendingPageJumps.delete(uri);
                        console.log(`期限切れのページジャンプ情報を削除しました（URI: ${uri}）`);
                    }
                }
                
                // ファイル名マップの古いエントリを削除
                for (const [fileName, info] of PdfEditorProvider.pendingPageJumpsByFileName.entries()) {
                    if (info.timestamp < expirationTime) {
                        PdfEditorProvider.pendingPageJumpsByFileName.delete(fileName);
                        console.log(`期限切れのページジャンプ情報を削除しました（ファイル名: ${fileName}）`);
                    }
                }
            }
        } catch (error) {
            vscode.window.showErrorMessage(`PDFファイルを開けませんでした: ${error}`);
        }
    }
}
/**
 * PDFファイルのカスタムエディタプロバイダ
 */
class PdfEditorProvider implements vscode.CustomReadonlyEditorProvider {
    // 現在開いているPDFビューアのパネルを保持
    private static panels = new Map<string, vscode.WebviewPanel>();
    
    // ページジャンプの保留情報を保持（URIをキーとするマップ）
    public static pendingPageJumps = new Map<string, { page: number, timestamp: number }>();
    
    // ファイル名をキーとするページジャンプ情報も保持（URIが一致しない場合のバックアップ）
    public static pendingPageJumpsByFileName = new Map<string, { page: number, timestamp: number }>();
    
    constructor(private readonly context: vscode.ExtensionContext) {
        // 特定のページに移動するコマンドを登録
        this.context.subscriptions.push(
            vscode.commands.registerCommand('contractmate.goToPage', (page: number, uri?: vscode.Uri) => {
                console.log(`contractmate.goToPage コマンドが実行されました: ページ ${page}`);
                
                // 開いているすべてのパネルを取得
                if (PdfEditorProvider.panels.size > 0) {
                    let targetPanel: vscode.WebviewPanel | undefined;
                    
                    // URIが指定されている場合は、そのURIに対応するパネルを探す
                    if (uri) {
                        const uriString = uri.toString();
                        console.log(`指定されたURI: ${uriString}`);
                        targetPanel = PdfEditorProvider.panels.get(uriString);
                    }
                    
                    // URIに対応するパネルが見つからない場合は、アクティブなエディタグループにあるパネルを探す
                    if (!targetPanel) {
                        const activeViewColumn = vscode.window.activeTextEditor?.viewColumn || vscode.ViewColumn.One;
                        console.log(`アクティブなエディタグループ: ${activeViewColumn}`);
                        
                        for (const [panelUri, panel] of PdfEditorProvider.panels.entries()) {
                            console.log(`パネルURI: ${panelUri}, ビューカラム: ${panel.viewColumn}`);
                            if (panel.viewColumn === activeViewColumn) {
                                targetPanel = panel;
                                break;
                            }
                        }
                    }
                    
                    // それでもパネルが見つからない場合は、最後に開いたパネルを使用
                    if (!targetPanel) {
                        console.log('最後に開いたパネルを使用します');
                        targetPanel = Array.from(PdfEditorProvider.panels.values())[PdfEditorProvider.panels.size - 1];
                    }
                    
                    // ページ移動メッセージを送信
                    if (targetPanel) {
                        console.log(`ページ ${page} に移動するメッセージを送信します`);
                        targetPanel.webview.postMessage({ command: 'goToPage', page });
                    } else {
                        console.log('対象のパネルが見つかりませんでした');
                    }
                } else {
                    console.log('開いているPDFパネルがありません');
                }
            })
        );
    }

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
        try {
            // パネルを保存
            PdfEditorProvider.panels.set(document.uri.toString(), webviewPanel);
            
            // パネルが閉じられたときの処理
            webviewPanel.onDidDispose(() => {
                PdfEditorProvider.panels.delete(document.uri.toString());
            });
            
            // WebViewの設定
            webviewPanel.webview.options = {
                enableScripts: true,
                localResourceRoots: [
                    vscode.Uri.joinPath(this.context.extensionUri, 'node_modules', 'pdfjs-dist'),
                    vscode.Uri.joinPath(this.context.extensionUri, 'media'),
                    vscode.Uri.file(path.dirname(document.uri.fsPath)) // PDFファイルのディレクトリを許可
                ]
            };

            // PDFファイルの内容を取得
            const pdfData = await this.getPdfData(document.uri);
            
            // WebViewのHTMLを設定
            webviewPanel.webview.html = this.getHtmlForWebview(webviewPanel.webview, document.uri, pdfData);
            
            // メッセージハンドラを設定
            webviewPanel.webview.onDidReceiveMessage(
                message => {
                    switch (message.command) {
                        case 'error':
                            vscode.window.showErrorMessage(`PDF表示エラー: ${message.text}`);
                            console.error('PDF表示エラー:', message.text);
                            break;
                        case 'log':
                            console.log('PDFビューアログ:', message.text);
                            break;
                        case 'ready':
                            // PDFビューアの準備ができたことを通知
                            console.log('PDFビューアの準備完了');
                            
                            // メッセージからPDFの総ページ数と現在のページを取得
                            const totalPages = message.totalPages || 1;
                            const currentPage = message.currentPage || 1;
                            console.log(`PDFの総ページ数: ${totalPages}, 現在のページ: ${currentPage}`);
                            
                            // 現在のURIとファイル名を取得
                            const originalUri = document.uri.toString();
                            const cleanUri = originalUri.replace(/[?#].*$/, '');
                            const currentFileName = path.basename(document.uri.fsPath);
                            
                            console.log(`現在のURI: ${originalUri}, クリーンURI: ${cleanUri}, ファイル名: ${currentFileName}`);
                            
                            // 複数のバリエーションでページジャンプ情報を検索
                            let pendingJump = null;
                            
                            // URIからページ番号を直接抽出（ハッシュフラグメントから）
                            let pageFromUri = null;
                            
                            // #page=N 形式のチェック
                            const pageMatch = originalUri.match(/#page=(\d+)/i);
                            if (pageMatch && pageMatch[1]) {
                                pageFromUri = parseInt(pageMatch[1], 10);
                                console.log(`URIのハッシュフラグメントからページ番号を直接抽出しました(#page=N形式): ${pageFromUri}`);
                            }
                            // #page%3DN 形式のチェック（URLエンコードされたハッシュフラグメント）
                            else {
                                const encodedPageMatch = originalUri.match(/#page%3D(\d+)/i);
                                if (encodedPageMatch && encodedPageMatch[1]) {
                                    pageFromUri = parseInt(encodedPageMatch[1], 10);
                                    console.log(`URIのハッシュフラグメントからページ番号を直接抽出しました(#page%3DN形式): ${pageFromUri}`);
                                }
                            }
                            
                            // URIから直接抽出したページ番号があれば、それを使用
                            if (pageFromUri !== null && pageFromUri > 0) {
                                // ページ番号が有効範囲内かチェック
                                let validPage = pageFromUri;
                                if (pageFromUri > totalPages) {
                                    console.log(`URIから抽出したページ番号 ${pageFromUri} がPDFの総ページ数 ${totalPages} を超えています。最終ページを表示します。`);
                                    validPage = totalPages;
                                } else if (pageFromUri < 1) {
                                    console.log(`URIから抽出したページ番号 ${pageFromUri} が無効です。最初のページを表示します。`);
                                    validPage = 1;
                                } else {
                                    console.log(`URIから抽出したページ番号 ${pageFromUri} に直接ジャンプします`);
                                }
                                
                                // ページジャンプを実行
                                webviewPanel.webview.postMessage({
                                    command: 'goToPage',
                                    page: validPage,
                                    scrollMode: 0, // 0: vertical, 1: horizontal, 2: wrapped
                                    forceJump: true
                                });
                                
                                // pendingJumpを設定して、後続の処理をスキップ
                                pendingJump = { page: pageFromUri, timestamp: Date.now() };
                            }
                            
                            // 1. 元のURI（変更なし）
                            pendingJump = PdfEditorProvider.pendingPageJumps.get(originalUri);
                            if (pendingJump) {
                                console.log(`元のURIに基づいてページジャンプ情報を見つけました: ${originalUri}`);
                            }
                            
                            // 2. クリーンURI（ハッシュとクエリパラメータを除去）
                            if (!pendingJump) {
                                pendingJump = PdfEditorProvider.pendingPageJumps.get(cleanUri);
                                if (pendingJump) {
                                    console.log(`クリーンURIに基づいてページジャンプ情報を見つけました: ${cleanUri}`);
                                }
                            }
                            
                            // 3. デコードされたURI
                            if (!pendingJump) {
                                try {
                                    const decodedUri = decodeURIComponent(cleanUri);
                                    if (decodedUri !== cleanUri) {
                                        pendingJump = PdfEditorProvider.pendingPageJumps.get(decodedUri);
                                        if (pendingJump) {
                                            console.log(`デコードされたURIに基づいてページジャンプ情報を見つけました: ${decodedUri}`);
                                        }
                                    }
                                } catch (e) {
                                    console.log(`URIのデコードに失敗: ${e}`);
                                }
                            }
                            
                            // 4. ファイル名
                            if (!pendingJump) {
                                pendingJump = PdfEditorProvider.pendingPageJumpsByFileName.get(currentFileName);
                                if (pendingJump) {
                                    console.log(`ファイル名に基づいてページジャンプ情報を見つけました: ${currentFileName}`);
                                }
                            }
                            
                            // 5. ファイルパスのみ（スキーム部分を除去）
                            if (!pendingJump) {
                                const filePathOnly = document.uri.fsPath;
                                pendingJump = PdfEditorProvider.pendingPageJumps.get(filePathOnly);
                                if (pendingJump) {
                                    console.log(`ファイルパスのみに基づいてページジャンプ情報を見つけました: ${filePathOnly}`);
                                }
                            }
                            
                            // 6. ファイルURIスキーム付き（file:///形式）
                            if (!pendingJump) {
                                const fileUri = `file:///${document.uri.fsPath.replace(/\\/g, '/')}`;
                                pendingJump = PdfEditorProvider.pendingPageJumps.get(fileUri);
                                if (pendingJump) {
                                    console.log(`ファイルURIスキーム付きに基づいてページジャンプ情報を見つけました: ${fileUri}`);
                                }
                            }
                            
                            // 7. クエリパラメータから直接ページ番号を抽出
                            if (!pendingJump && originalUri.includes('?')) {
                                const queryMatch = originalUri.match(/[?&]page=(\d+)/i);
                                if (queryMatch && queryMatch[1]) {
                                    const pageFromQuery = parseInt(queryMatch[1], 10);
                                    console.log(`クエリパラメータから直接ページ番号を抽出しました: ${pageFromQuery}`);
                                    pendingJump = { page: pageFromQuery, timestamp: Date.now() };
                                }
                            }
                            
                            // 8. ハッシュフラグメントから直接ページ番号を抽出
                            if (!pendingJump && originalUri.includes('#')) {
                                // #page=N 形式のチェック
                                const pageMatch = originalUri.match(/#page=(\d+)/i);
                                if (pageMatch && pageMatch[1]) {
                                    const pageFromHash = parseInt(pageMatch[1], 10);
                                    console.log(`ハッシュフラグメントから直接ページ番号を抽出しました(#page=N形式): ${pageFromHash}`);
                                    pendingJump = { page: pageFromHash, timestamp: Date.now() };
                                }
                                // #N 形式のチェック
                                else {
                                    const hashMatch = originalUri.match(/#(\d+)$/);
                                    if (hashMatch && hashMatch[1]) {
                                        const pageFromHash = parseInt(hashMatch[1], 10);
                                        console.log(`ハッシュフラグメントから直接ページ番号を抽出しました(#N形式): ${pageFromHash}`);
                                        pendingJump = { page: pageFromHash, timestamp: Date.now() };
                                    }
                                }
                            }
                            
                            // 保留中のページジャンプがあれば実行（URIから直接抽出した場合は既に実行済み）
                            if (pendingJump && pageFromUri === null) {
                                const { page, timestamp } = pendingJump;
                                
                                // ページジャンプ情報が古すぎないか確認（5分以内）
                                const now = Date.now();
                                if (now - timestamp > 5 * 60 * 1000) {
                                    console.log(`ページジャンプ情報が古すぎるためスキップします`);
                                } else {
                                    // ページ番号が有効範囲内かチェック
                                    let validPage = page;
                                    if (page > totalPages) {
                                        validPage = totalPages;
                                    } else if (page < 1) {
                                        validPage = 1;
                                    }
                                    
                                    console.log(`保存されていたページジャンプ情報に基づいてページ ${validPage} に移動します`);
                                    
                                    // ページジャンプを実行
                                    webviewPanel.webview.postMessage({
                                        command: 'goToPage',
                                        page: validPage,
                                        scrollMode: 0,
                                        forceJump: true
                                    });
                                    
                                    // 保留情報をクリア
                                    PdfEditorProvider.pendingPageJumps.delete(originalUri);
                                    PdfEditorProvider.pendingPageJumps.delete(cleanUri);
                                    PdfEditorProvider.pendingPageJumpsByFileName.delete(currentFileName);
                                }
                            }
                            break;
                    }
                }
            );
        } catch (error) {
            vscode.window.showErrorMessage(`PDFビューアの初期化に失敗しました: ${error}`);
            console.error('PDFビューア初期化エラー:', error);
        }
    }

    /**
     * PDFファイルのデータを取得する
     */
    private async getPdfData(uri: vscode.Uri): Promise<Uint8Array> {
        try {
            return await vscode.workspace.fs.readFile(uri);
        } catch (error) {
            console.error('PDFファイル読み込みエラー:', error);
            throw new Error(`PDFファイルの読み込みに失敗しました: ${error}`);
        }
    }

    /**
     * WebViewのHTMLを生成する
     */
    private getHtmlForWebview(webview: vscode.Webview, uri: vscode.Uri, pdfData: Uint8Array): string {
        // PDF.jsのパスを取得
        const pdfJsDistPath = vscode.Uri.joinPath(this.context.extensionUri, 'node_modules', 'pdfjs-dist');
        const pdfJsPath = webview.asWebviewUri(vscode.Uri.joinPath(pdfJsDistPath, 'build', 'pdf.mjs')).toString();
        const pdfWorkerPath = webview.asWebviewUri(vscode.Uri.joinPath(pdfJsDistPath, 'build', 'pdf.worker.mjs')).toString();
        const pdfViewerCssPath = webview.asWebviewUri(vscode.Uri.joinPath(pdfJsDistPath, 'web', 'pdf_viewer.css')).toString();
        
        // cMapUrlを事前に計算（末尾にスラッシュを追加）
        const cMapUrl = webview.asWebviewUri(vscode.Uri.joinPath(this.context.extensionUri, 'node_modules', 'pdfjs-dist', 'cmaps')).toString() + '/';
        
        // PDFデータをBase64エンコード
        const pdfDataBase64 = Buffer.from(pdfData).toString('base64');
        
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
                        padding: 0;
                        background-color: var(--vscode-editor-background);
                        color: var(--vscode-editor-foreground);
                        font-family: var(--vscode-font-family);
                        font-size: var(--vscode-font-size);
                        overflow: hidden;
                    }
                    #toolbar {
                        position: fixed;
                        top: 0;
                        left: 0;
                        right: 0;
                        height: 40px;
                        background-color: var(--vscode-editor-background);
                        border-bottom: 1px solid var(--vscode-panel-border);
                        display: flex;
                        align-items: center;
                        padding: 0 10px;
                        z-index: 100;
                    }
                    #viewerContainer {
                        position: absolute;
                        top: 40px;
                        left: 0;
                        right: 0;
                        bottom: 0;
                        overflow: auto;
                    }
                    #viewer {
                        position: relative;
                        width: 100%;
                        min-height: 100%;
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: flex-start; /* 上から下への順序を明示的に指定 */
                    }
                    .page {
                        margin-bottom: 10px;
                        box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2);
                        background-color: white;
                    }
                    .loadingMessage {
                        color: var(--vscode-editor-foreground);
                        text-align: center;
                        padding-top: 50px;
                        font-size: 16px;
                    }
                    .toolbarButton {
                        background-color: var(--vscode-button-background);
                        color: var(--vscode-button-foreground);
                        border: none;
                        padding: 5px 10px;
                        margin: 0 5px;
                        cursor: pointer;
                        border-radius: 2px;
                    }
                    .toolbarButton:hover {
                        background-color: var(--vscode-button-hoverBackground);
                    }
                    .toolbarButton.active {
                        background-color: var(--vscode-button-hoverBackground);
                        font-weight: bold;
                    }
                    #pageInfo {
                        margin: 0 10px;
                        color: var(--vscode-editor-foreground);
                    }
                    #pageInput {
                        width: 40px;
                        background-color: var(--vscode-input-background);
                        color: var(--vscode-input-foreground);
                        border: 1px solid var(--vscode-input-border);
                        padding: 2px 5px;
                    }
                </style>
            </head>
            <body>
                <div id="toolbar">
                    <button id="prevPage" class="toolbarButton">前のページ</button>
                    <button id="nextPage" class="toolbarButton">次のページ</button>
                    <span id="pageInfo">
                        <input id="pageInput" type="number" min="1" value="1"> / <span id="pageCount">0</span>
                    </span>
                    <button id="zoomOut" class="toolbarButton">縮小</button>
                    <button id="zoomIn" class="toolbarButton">拡大</button>
                    <button id="fitWidth" class="toolbarButton">幅に合わせる</button>
                    <button id="resetZoom" class="toolbarButton">リセット</button>
                </div>
                <div id="viewerContainer">
                    <div id="viewer"></div>
                </div>
                <div id="loadingMessage" class="loadingMessage">PDFを読み込んでいます...</div>

                <script type="module">
                    // PDF.jsをインポート
                    import * as pdfjsLib from "${pdfJsPath}";
                    
                    // VSCodeのWebViewとの通信用
                    const vscode = acquireVsCodeApi();
                    
                    // デバッグ用のログ関数（重要なログのみ出力）
                    function logToVSCode(message, isImportant = false) {
                        // 重要なログのみ出力する
                        if (isImportant) {
                            vscode.postMessage({
                                command: 'log',
                                text: message
                            });
                        }
                    }
                    
                    // エラー報告用の関数
                    function reportErrorToVSCode(error) {
                        vscode.postMessage({
                            command: 'error',
                            text: error.toString()
                        });
                    }
                    
                    try {
                        // PDF.jsのワーカーパスを設定
                        pdfjsLib.GlobalWorkerOptions.workerSrc = "${pdfWorkerPath}";
                        logToVSCode('PDF.jsワーカーパス設定完了');
                        
                        // 変数の初期化
                        let pdfDoc = null;
                        let pageNum = 1;
                        let scale = 1.5;
                        let pagesCache = {};
                        let containerWidth = document.getElementById('viewerContainer').clientWidth - 40;
                        let isRendering = false;
                        let renderQueue = [];
                        
                        // VSCodeからのメッセージを受け取る
                        window.addEventListener('message', event => {
                            const message = event.data;
                            switch (message.command) {
                                case 'goToPage':
                                    if (pdfDoc && message.page > 0) {
                                        // ページ番号がPDFの総ページ数を超える場合は、最終ページを表示
                                        if (message.page > pdfDoc.numPages) {
                                            pageNum = pdfDoc.numPages;
                                        } else {
                                            pageNum = message.page;
                                        }
                                        
                                        // ページ入力フィールドを更新
                                        pageInput.value = pageNum;
                                        
                                        // スクロールモードを設定（メッセージから取得、デフォルトは垂直スクロール）
                                        const scrollMode = message.scrollMode !== undefined ? message.scrollMode : 0;
                                        
                                        try {
                                            // 強制ジャンプフラグを確認
                                            const forceJump = message.forceJump === true;
                                            
                                            // 1. 既存のページをクリア
                                            viewer.innerHTML = '';
                                            renderQueue = [];
                                            isRendering = false;
                                            pagesCache = {};
                                            
                                            // 2. 表示モードを設定（スクロールモードに応じて）
                                            // スクロールモードに応じてビューアのスタイルを変更
                                            if (scrollMode === 0) { // 垂直スクロール
                                                viewer.style.flexDirection = 'column';
                                            } else if (scrollMode === 1) { // 水平スクロール
                                                viewer.style.flexDirection = 'row';
                                            } else if (scrollMode === 2) { // ラップモード
                                                viewer.style.flexDirection = 'row';
                                                viewer.style.flexWrap = 'wrap';
                                            }
                                            
                                            // 3. 指定されたページを取得
                                            pdfDoc.getPage(pageNum).then(page => {
                                                // 4. 現在のページを最優先でレンダリング
                                                queueRenderPage(pageNum);
                                                
                                                // 5. 前後のページも追加
                                                if (pageNum > 1) {
                                                    queueRenderPage(pageNum - 1);
                                                }
                                                if (pageNum < pdfDoc.numPages) {
                                                    queueRenderPage(pageNum + 1);
                                                }
                                                
                                                // 6. 強化されたページジャンプ処理
                                                const enhancedJumpToPage = (attempts = 0) => {
                                                    // 指定されたページの要素を取得
                                                    const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                                    
                                                    if (targetElement) {
                                                        
                                                        try {
                                                            // スクロール位置を設定（ページの上部に合わせる）
                                                            const scrollPosition = targetElement.offsetTop - 40;
                                                            
                                                            // 即時スクロール（スムーズスクロールなし）
                                                            viewerContainer.scrollTop = scrollPosition;
                                                            
                                                            // スクロール完了を確認するために少し待機
                                                            setTimeout(() => {
                                                                // 現在のスクロール位置を確認
                                                                const currentScrollTop = viewerContainer.scrollTop;
                                                                
                                                                // スクロール位置が期待通りか確認
                                                                const expectedScrollTop = targetElement.offsetTop - 40;
                                                                const scrollDiff = Math.abs(currentScrollTop - expectedScrollTop);
                                                                
                                                                if (scrollDiff > 50) {
                                                                    viewerContainer.scrollTop = expectedScrollTop;
                                                                }
                                                                
                                                                // スクロール完了を通知
                                                                vscode.postMessage({
                                                                    command: 'pageJumpComplete',
                                                                    page: pageNum,
                                                                    success: true
                                                                });
                                                            }, 500);
                                                        } catch (error) {
                                                            // エラーが発生した場合でも再試行
                                                            if (attempts < 10) {
                                                                setTimeout(() => enhancedJumpToPage(attempts + 1), 300);
                                                            } else {
                                                                // 失敗を通知
                                                                vscode.postMessage({
                                                                    command: 'pageJumpComplete',
                                                                    page: pageNum,
                                                                    success: false,
                                                                    error: error.toString()
                                                                });
                                                            }
                                                        }
                                                    } else if (attempts < 30) { // 試行回数を増やす
                                                        // ページがまだレンダリングされていない可能性があるため、再試行
                                                        setTimeout(() => enhancedJumpToPage(attempts + 1), 200);
                                                    } else {
                                                        
                                                        // 代替方法: すべてのページを再レンダリングして、スクロール位置を計算
                                                        try {
                                                            // ページの高さを推定（最初のページの高さを使用）
                                                            const firstPage = document.querySelector('.page');
                                                            if (firstPage) {
                                                                const pageHeight = firstPage.clientHeight;
                                                                const estimatedPosition = (pageNum - 1) * (pageHeight + 10); // 10はページ間のマージン
                                                                viewerContainer.scrollTop = estimatedPosition;
                                                            }
                                                        } catch (error) {
                                                            console.error('ページジャンプ失敗:', error);
                                                        }
                                                    }
                                                };
                                                
                                                // ページジャンプを実行
                                                enhancedJumpToPage();
                                            }).catch(error => {
                                                reportErrorToVSCode(error);
                                            });
                                        } catch (error) {
                                            reportErrorToVSCode(error);
                                        }
                                    }
                                    break;
                            }
                        });
                        
                        // UI要素の取得
                        const loadingMessage = document.getElementById('loadingMessage');
                        const viewer = document.getElementById('viewer');
                        const viewerContainer = document.getElementById('viewerContainer');
                        const prevButton = document.getElementById('prevPage');
                        const nextButton = document.getElementById('nextPage');
                        const pageInput = document.getElementById('pageInput');
                        const pageCount = document.getElementById('pageCount');
                        const zoomInButton = document.getElementById('zoomIn');
                        const zoomOutButton = document.getElementById('zoomOut');
                        const fitWidthButton = document.getElementById('fitWidth');
                        const resetZoomButton = document.getElementById('resetZoom');
                        
                        // ページをレンダリングするキュー処理
                        function queueRenderPage(num) {
                            if (renderQueue.indexOf(num) === -1) {
                                renderQueue.push(num);
                                processRenderQueue();
                            }
                        }
                        
                        // レンダリングキューを処理
                        function processRenderQueue() {
                            if (isRendering || renderQueue.length === 0) {
                                return;
                            }
                            
                            isRendering = true;
                            const num = renderQueue.shift();
                            
                            // 既にレンダリング済みのページはスキップ
                            if (document.querySelector('.page[data-page-number="' + num + '"]')) {
                                isRendering = false;
                                processRenderQueue();
                                return;
                            }
                            
                            // ページをレンダリング
                            pdfDoc.getPage(num).then(function(page) {
                                const viewport = page.getViewport({ scale: scale });
                                
                                // ページ要素を作成
                                const pageDiv = document.createElement('div');
                                pageDiv.className = 'page';
                                pageDiv.setAttribute('data-page-number', num);
                                
                                // キャンバスを作成
                                const canvas = document.createElement('canvas');
                                const context = canvas.getContext('2d');
                                canvas.height = viewport.height;
                                canvas.width = viewport.width;
                                
                                // ページをキャンバスにレンダリング
                                const renderContext = {
                                    canvasContext: context,
                                    viewport: viewport
                                };
                                
                                pageDiv.appendChild(canvas);
                                
                                // ページを適切な位置に挿入
                                let inserted = false;
                                const pages = viewer.querySelectorAll('.page');
                                for (let i = 0; i < pages.length; i++) {
                                    const pageNumber = parseInt(pages[i].getAttribute('data-page-number'));
                                    if (pageNumber > num) {
                                        viewer.insertBefore(pageDiv, pages[i]);
                                        inserted = true;
                                        break;
                                    }
                                }
                                
                                if (!inserted) {
                                    viewer.appendChild(pageDiv);
                                }
                                
                                // ページをレンダリング
                                page.render(renderContext).promise.then(function() {
                                    // レンダリング完了
                                    isRendering = false;
                                    
                                    // キャッシュに保存
                                    pagesCache[num] = true;
                                    
                                    // 次のページをレンダリング
                                    processRenderQueue();
                                    
                                    // ローディングメッセージを非表示
                                    loadingMessage.style.display = 'none';
                                }).catch(function(error) {
                                    console.error('ページレンダリングエラー:', error);
                                    isRendering = false;
                                    processRenderQueue();
                                });
                            }).catch(function(error) {
                                console.error('ページ取得エラー:', error);
                                isRendering = false;
                                processRenderQueue();
                            });
                        }
                        
                        // ページを更新する関数
                        function renderPage(num) {
                            // ページ番号を更新
                            pageNum = num;
                            pageInput.value = pageNum;
                            
                            // 前後のページをキューに追加
                            queueRenderPage(num);
                            
                            // 前後のページも事前にレンダリング
                            if (num > 1) {
                                queueRenderPage(num - 1);
                            }
                            if (num < pdfDoc.numPages) {
                                queueRenderPage(num + 1);
                            }
                            
                            // 表示されているページにスクロール
                            const targetPage = document.querySelector('.page[data-page-number="' + num + '"]');
                            if (targetPage) {
                                viewerContainer.scrollTop = targetPage.offsetTop - 40;
                            }
                            
                            return Promise.resolve();
                        }
                        
                        // Base64エンコードされたPDFデータをデコード
                        const pdfData = atob('${pdfDataBase64}');
                        logToVSCode('PDFデータのデコード完了: ' + pdfData.length + ' バイト', true);
                        
                        // Uint8Arrayに変換
                        const uint8Array = new Uint8Array(pdfData.length);
                        for (let i = 0; i < pdfData.length; i++) {
                            uint8Array[i] = pdfData.charCodeAt(i);
                        }
                        logToVSCode('Uint8Arrayへの変換完了: ' + uint8Array.length + ' バイト', true);
                        
                        // PDFを読み込む
                        logToVSCode('PDFの読み込み開始', true);
                        
                        try {
                            // PDF.jsのオプションを設定（シンプル化）
                            const pdfOptions = {
                                data: uint8Array,
                                // 最小限のオプションのみ設定
                                cMapUrl: '${cMapUrl}',
                                cMapPacked: true
                            };
                            
                            logToVSCode('PDF.jsオプション設定完了', true);
                            
                            // PDFの読み込みとレンダリング
                            pdfjsLib.getDocument(pdfOptions).promise
                                .then(function(pdf) {
                                    logToVSCode('PDFの読み込み成功: ' + pdf.numPages + ' ページ', true);
                                    pdfDoc = pdf;
                                    loadingMessage.style.display = 'none';
                                    pageCount.textContent = pdf.numPages;
                                    pageInput.max = pdf.numPages;
                                    
                                    // 最初のページをレンダリング（直接レンダリング）
                                    renderPage(pageNum).then(() => {
                                        logToVSCode('最初のページのレンダリング完了', true);
                                    }).catch(error => {
                                        logToVSCode('最初のページのレンダリングエラー: ' + error, true);
                                    });
                                    
                                    // PDFビューアの準備完了を通知
                                    vscode.postMessage({
                                        command: 'ready',
                                        totalPages: pdf.numPages,
                                        currentPage: pageNum
                                    });
                                })
                                .catch(function(error) {
                                    loadingMessage.textContent = 'PDFの読み込みに失敗しました: ' + error;
                                    reportErrorToVSCode(error);
                                    logToVSCode('PDFの読み込みエラー: ' + error, true);
                                });
                        } catch (error) {
                            loadingMessage.textContent = 'PDFの初期化に失敗しました: ' + error;
                            reportErrorToVSCode(error);
                            logToVSCode('PDF初期化エラー: ' + error, true);
                        }
                        
                        // イベントリスナーの設定
                        prevButton.addEventListener('click', function() {
                            if (pageNum <= 1) {
                                return;
                            }
                            pageNum--;
                            renderPage(pageNum);
                        });
                        
                        nextButton.addEventListener('click', function() {
                            if (pageNum >= pdfDoc.numPages) {
                                return;
                            }
                            pageNum++;
                            renderPage(pageNum);
                        });
                        
                        pageInput.addEventListener('change', function() {
                            const num = parseInt(pageInput.value);
                            if (num > 0 && num <= pdfDoc.numPages) {
                                pageNum = num;
                                renderPage(pageNum);
                            } else {
                                pageInput.value = pageNum;
                            }
                        });
                        
                        zoomInButton.addEventListener('click', function() {
                            scale *= 1.2;
                            viewer.innerHTML = '';
                            renderPage(pageNum);
                        });
                        
                        zoomOutButton.addEventListener('click', function() {
                            scale /= 1.2;
                            viewer.innerHTML = '';
                            renderPage(pageNum);
                        });
                        
                        fitWidthButton.addEventListener('click', function() {
                            // コンテナの幅に合わせてスケールを計算
                            pdfDoc.getPage(pageNum).then(function(page) {
                                const viewport = page.getViewport({ scale: 1.0 });
                                scale = containerWidth / viewport.width;
                                viewer.innerHTML = '';
                                renderPage(pageNum);
                            });
                        });
                        
                        resetZoomButton.addEventListener('click', function() {
                            scale = 1.5;
                            viewer.innerHTML = '';
                            renderPage(pageNum);
                        });
                        
                        // ウィンドウサイズ変更時の処理
                        window.addEventListener('resize', function() {
                            containerWidth = document.getElementById('viewerContainer').clientWidth - 40;
                        });
                        
                        // スクロール時の処理（ビューポート内のページを動的にレンダリング）
                        viewerContainer.addEventListener('scroll', function() {
                            if (!pdfDoc) return;
                            
                            const visibleTop = viewerContainer.scrollTop;
                            const visibleBottom = visibleTop + viewerContainer.clientHeight;
                            
                            // 表示範囲内のページを特定
                            const pages = viewer.querySelectorAll('.page');
                            let visiblePageNum = pageNum;
                            
                            for (let i = 0; i < pages.length; i++) {
                                const page = pages[i];
                                const pageTop = page.offsetTop;
                                const pageBottom = pageTop + page.clientHeight;
                                
                                // ページが表示範囲内にあるか確認
                                if (pageTop < visibleBottom && pageBottom > visibleTop) {
                                    const num = parseInt(page.getAttribute('data-page-number'));
                                    visiblePageNum = num;
                                    break;
                                }
                            }
                            
                            // 現在のページ番号を更新
                            if (visiblePageNum !== pageNum) {
                                pageNum = visiblePageNum;
                                pageInput.value = pageNum;
                            }
                            
                            // 表示範囲の前後のページをレンダリング
                            for (let i = 1; i <= pdfDoc.numPages; i++) {
                                const page = document.querySelector('.page[data-page-number="' + i + '"]');
                                if (!page) continue;
                                
                                const pageTop = page.offsetTop;
                                const pageBottom = pageTop + page.clientHeight;
                                
                                // 表示範囲の前後のページをレンダリング
                                if (pageBottom > visibleTop - 1000 && pageTop < visibleBottom + 1000) {
                                    queueRenderPage(i);
                                }
                            }
                        });
                    } catch (error) {
                        loadingMessage.textContent = 'PDFの初期化に失敗しました: ' + error;
                        reportErrorToVSCode(error);
                        logToVSCode('PDF初期化エラー: ' + error, true);
                    }
                </script>
            </body>
            </html>
        `;
    }
}