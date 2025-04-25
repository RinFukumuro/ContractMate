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
                                                            
                                                            // スムーズスクロールを使用
                                                            viewerContainer.scrollTo({
                                                                top: scrollPosition,
                                                                behavior: 'smooth'
                                                            });
                                                            
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
                                                                const estimatedPosition = (pageNum - 1) * (pageHeight + 10); // 10はマージン
                                                                
                                                                logToVSCode('ページの高さを推定: ' + pageHeight + 'px');
                                                                logToVSCode('推定スクロール位置: ' + estimatedPosition + 'px');
                                                                
                                                                // 推定位置にスクロール
                                                                viewerContainer.scrollTop = estimatedPosition;
                                                                
                                                                // スクロール完了を通知
                                                                vscode.postMessage({
                                                                    command: 'pageJumpComplete',
                                                                    page: pageNum,
                                                                    success: true,
                                                                    warning: '推定位置にスクロールしました'
                                                                });
                                                            } else {
                                                                logToVSCode('ページ要素が見つかりません。スクロールできません。');
                                                                
                                                                // 失敗を通知
                                                                vscode.postMessage({
                                                                    command: 'pageJumpComplete',
                                                                    page: pageNum,
                                                                    success: false,
                                                                    error: 'ページ要素が見つかりません'
                                                                });
                                                            }
                                                        } catch (error) {
                                                            logToVSCode('代替スクロール方法でエラーが発生しました: ' + error);
                                                            
                                                            // 失敗を通知
                                                            vscode.postMessage({
                                                                command: 'pageJumpComplete',
                                                                page: pageNum,
                                                                success: false,
                                                                error: error.toString()
                                                            });
                                                        }
                                                    }
                                                };
                                                
                                                // レンダリング待機を開始
                                                setTimeout(() => enhancedJumpToPage(), 300);
                                            }).catch(error => {
                                                logToVSCode('ページ取得エラー: ' + error);
                                                
                                                // エラーを通知
                                                vscode.postMessage({
                                                    command: 'pageJumpComplete',
                                                    page: pageNum,
                                                    success: false,
                                                    error: error.toString()
                                                });
                                            });
                                        } catch (error) {
                                            logToVSCode('ページジャンプ処理でエラーが発生しました: ' + error);
                                            
                                            // エラーを通知
                                            vscode.postMessage({
                                                command: 'pageJumpComplete',
                                                page: pageNum,
                                                success: false,
                                                error: error.toString()
                                            });
                                        }
                                    }
                                    break;
                            }
                        });

                        // UI要素
                        const viewerContainer = document.getElementById('viewerContainer');
                        const viewer = document.getElementById('viewer');
                        const loadingMessage = document.getElementById('loadingMessage');
                        const prevButton = document.getElementById('prevPage');
                        const nextButton = document.getElementById('nextPage');
                        const pageInput = document.getElementById('pageInput');
                        const pageCount = document.getElementById('pageCount');
                        const zoomInButton = document.getElementById('zoomIn');
                        const zoomOutButton = document.getElementById('zoomOut');
                        const fitWidthButton = document.getElementById('fitWidth');
                        const resetZoomButton = document.getElementById('resetZoom');
                        
                        logToVSCode('UI要素の初期化完了');
                        
                        // スクロールタイマー用の変数
                        let scrollTimer = null;
                        
                        // スクロールイベントを監視して、現在表示されているページを認識する
                        viewerContainer.addEventListener('scroll', function() {
                            if (!pdfDoc || isRendering) return;
                            
                            // スクロール中のタイマーをクリア
                            if (scrollTimer) {
                                clearTimeout(scrollTimer);
                            }
                            
                            // スクロールが停止した後に現在のページを更新（デバウンス処理）
                            scrollTimer = setTimeout(() => {
                                // スクロール位置を取得
                                const scrollTop = viewerContainer.scrollTop;
                                const viewportHeight = viewerContainer.clientHeight;
                                
                                // 表示されているすべてのページ要素を取得
                                const pageElements = document.querySelectorAll('.page');
                                if (pageElements.length === 0) return;
                                
                                // 各ページの可視率を計算し、最も可視率が高いページを特定
                                let mostVisiblePage = null;
                                let highestVisibleRatio = 0;
                                
                                pageElements.forEach(pageElement => {
                                    const rect = pageElement.getBoundingClientRect();
                                    
                                    // ページの上端と下端の位置（ビューポート相対）
                                    const pageTop = rect.top;
                                    const pageBottom = rect.bottom;
                                    
                                    // ビューポートの上端と下端の位置
                                    const viewportTop = 0;
                                    const viewportBottom = viewportHeight;
                                    
                                    // ページとビューポートの重なり部分を計算
                                    const overlapTop = Math.max(pageTop, viewportTop);
                                    const overlapBottom = Math.min(pageBottom, viewportBottom);
                                    
                                    // 重なりがある場合のみ計算
                                    if (overlapBottom > overlapTop) {
                                        // 重なっている高さ
                                        const overlapHeight = overlapBottom - overlapTop;
                                        // ページの高さに対する可視部分の割合
                                        const visibleRatio = overlapHeight / rect.height;
                                        
                                        // 可視率が30%以上かつ、これまでの最高可視率より高い場合
                                        if (visibleRatio > highestVisibleRatio && visibleRatio >= 0.3) {
                                            highestVisibleRatio = visibleRatio;
                                            mostVisiblePage = pageElement;
                                        }
                                    }
                                });
                                
                                // 可視率が30%以上のページが見つからなかった場合、最も可視率が高いページを選択
                                if (!mostVisiblePage && pageElements.length > 0) {
                                    // 各ページの可視率を再計算
                                    let bestPage = null;
                                    let bestRatio = 0;
                                    
                                    pageElements.forEach(pageElement => {
                                        const rect = pageElement.getBoundingClientRect();
                                        
                                        // ページの上端と下端の位置（ビューポート相対）
                                        const pageTop = rect.top;
                                        const pageBottom = rect.bottom;
                                        
                                        // ビューポートの上端と下端の位置
                                        const viewportTop = 0;
                                        const viewportBottom = viewportHeight;
                                        
                                        // ページとビューポートの重なり部分を計算
                                        const overlapTop = Math.max(pageTop, viewportTop);
                                        const overlapBottom = Math.min(pageBottom, viewportBottom);
                                        
                                        // 重なりがある場合のみ計算
                                        if (overlapBottom > overlapTop) {
                                            // 重なっている高さ
                                            const overlapHeight = overlapBottom - overlapTop;
                                            // ページの高さに対する可視部分の割合
                                            const visibleRatio = overlapHeight / rect.height;
                                            
                                            if (visibleRatio > bestRatio) {
                                                bestRatio = visibleRatio;
                                                bestPage = pageElement;
                                            }
                                        }
                                    });
                                    
                                    if (bestPage) {
                                        mostVisiblePage = bestPage;
                                        highestVisibleRatio = bestRatio;
                                    }
                                }
                                
                                // 最も可視率が高いページが見つかった場合
                                if (mostVisiblePage) {
                                    const newPageNum = parseInt(mostVisiblePage.dataset.pageNumber, 10);
                                    if (newPageNum !== pageNum) {
                                        pageNum = newPageNum;
                                        pageInput.value = pageNum;
                                        logToVSCode('スクロールにより現在のページを ' + pageNum + ' に更新しました（可視率: ' + Math.round(highestVisibleRatio * 100) + '%）');
                                    }
                                }
                            }, 100); // 100ms後に実行
                        });

                        // Base64エンコードされたPDFデータをデコード
                        const pdfData = atob('${pdfDataBase64}');
                        logToVSCode('PDFデータのデコード完了: ' + pdfData.length + ' バイト');
                        
                        // Uint8Arrayに変換
                        const uint8Array = new Uint8Array(pdfData.length);
                        for (let i = 0; i < pdfData.length; i++) {
                            uint8Array[i] = pdfData.charCodeAt(i);
                        }
                        logToVSCode('Uint8Arrayへの変換完了');
                        
                        // URLからハッシュフラグメントとクエリパラメータを抽出（ページ番号を取得するため）
                        const urlHash = window.location.hash;
                        const urlSearch = window.location.search;
                        let initialPage = 1;
                        
                        // デバッグ情報を出力
                        logToVSCode('URLハッシュ: ' + urlHash);
                        logToVSCode('URLクエリパラメータ: ' + urlSearch);
                        
                        // ページ内リンクのサポート
                        // ページ内の <a href="#page=N"> リンクをクリックしたときにページジャンプするように設定
                        document.addEventListener('click', function(e) {
                            // クリックされた要素がリンクかどうかを確認
                            if (e.target.tagName === 'A') {
                                const href = e.target.getAttribute('href');
                                if (href && (href.match(/#page=\d+/) || href.match(/#\d+$/))) {
                                    logToVSCode('ページ内リンクがクリックされました: ' + href);
                                    
                                    // デフォルトの動作を防止（ページのリロードを防ぐ）
                                    e.preventDefault();
                                    
                                    // URLハッシュを変更（hashchangeイベントが発火する）
                                    window.location.hash = href.substring(href.indexOf('#'));
                                }
                            }
                        });
                        
                        // 1. ハッシュフラグメントからページ番号を抽出
                        if (urlHash) {
                            // URLをデコード
                            let decodedHash;
                            try {
                                decodedHash = decodeURIComponent(urlHash);
                                logToVSCode('デコード後のURLハッシュ: ' + decodedHash);
                            } catch (e) {
                                logToVSCode('URLハッシュのデコードに失敗: ' + e);
                                decodedHash = urlHash;
                            }
                            
                            // #page=N 形式のチェック（デコード後）
                            const pageMatch = decodedHash.match(/#page=(\d+)/i);
                            if (pageMatch && pageMatch[1]) {
                                initialPage = parseInt(pageMatch[1], 10);
                                logToVSCode('デコードされたURLハッシュから抽出したページ番号(#page=N形式): ' + initialPage);
                            }
                            // #page%3DN 形式のチェック（デコード前）
                            else {
                                const encodedPageMatch = urlHash.match(/#page%3D(\d+)/i);
                                if (encodedPageMatch && encodedPageMatch[1]) {
                                    initialPage = parseInt(encodedPageMatch[1], 10);
                                    logToVSCode('エンコードされたURLハッシュから抽出したページ番号(#page%3DN形式): ' + initialPage);
                                }
                                // #N 形式のチェック（デコード後）
                                else {
                                    const hashMatch = decodedHash.match(/#(\d+)$/);
                                    if (hashMatch && hashMatch[1]) {
                                        initialPage = parseInt(hashMatch[1], 10);
                                        logToVSCode('デコードされたURLハッシュから抽出したページ番号(#N形式): ' + initialPage);
                                    }
                                }
                            }
                        }
                        
                        // 2. クエリパラメータからページ番号を抽出
                        if (initialPage === 1 && urlSearch) {
                            // URLをデコード
                            let decodedSearch;
                            try {
                                decodedSearch = decodeURIComponent(urlSearch);
                                logToVSCode('デコード後のURLクエリパラメータ: ' + decodedSearch);
                            } catch (e) {
                                logToVSCode('URLクエリパラメータのデコードに失敗: ' + e);
                                decodedSearch = urlSearch;
                            }
                            
                            // ?page=N 形式のチェック（デコード後）
                            const pageMatch = decodedSearch.match(/[?&]page=(\d+)/i);
                            if (pageMatch && pageMatch[1]) {
                                initialPage = parseInt(pageMatch[1], 10);
                                logToVSCode('デコードされたURLクエリパラメータから抽出したページ番号(?page=N形式): ' + initialPage);
                            }
                            // ?page%3DN 形式のチェック（デコード前）
                            else {
                                const encodedPageMatch = urlSearch.match(/[?&]page%3D(\d+)/i);
                                if (encodedPageMatch && encodedPageMatch[1]) {
                                    initialPage = parseInt(encodedPageMatch[1], 10);
                                    logToVSCode('エンコードされたURLクエリパラメータから抽出したページ番号(?page%3DN形式): ' + initialPage);
                                }
                            }
                        }
                        
                        // 初期ページ番号を設定
                        if (initialPage > 1) {
                            pageNum = initialPage;
                            logToVSCode('初期ページ番号を設定: ' + pageNum);
                        }

                        // PDFを読み込む
                        logToVSCode('PDFの読み込み開始');
                        
                        // PDF.jsのオプションを設定
                        const pdfOptions = {
                            data: uint8Array,
                            // PDF.jsの標準機能を有効化
                            enableXfa: true,
                            cMapUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.4.120/cmaps/',
                            cMapPacked: true,
                            // 追加のオプション
                            disableAutoFetch: false,
                            disableStream: false,
                            // レンダリングの品質を向上
                            useSystemFonts: true,
                            isEvalSupported: true,
                            standardFontDataUrl: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@3.4.120/standard_fonts/'
                        };
                        
                        logToVSCode('PDF.jsオプション設定完了: ' + JSON.stringify(pdfOptions));
                        
                        pdfjsLib.getDocument(pdfOptions).promise
                            .then(function(pdf) {
                                logToVSCode('PDFの読み込み成功: ' + pdf.numPages + ' ページ', true);
                                pdfDoc = pdf;
                                loadingMessage.style.display = 'none';
                                pageCount.textContent = pdf.numPages;
                                pageInput.max = pdf.numPages;
                                
                                // 最初のページをレンダリング
                                queueRenderPage(pageNum);
                                
                                // ビューポートの変更を監視
                                const resizeObserver = new ResizeObserver(entries => {
                                    for (let entry of entries) {
                                        if (entry.target === viewerContainer) {
                                            containerWidth = entry.contentRect.width - 40;
                                            if (fitWidthButton.classList.contains('active')) {
                                                fitToWidth();
                                            }
                                        }
                                    }
                                });
                                resizeObserver.observe(viewerContainer);
                                
                                // 初期状態で幅に合わせる
                                setTimeout(() => {
                                    fitToWidth();
                                    
                                    // ビューアのスタイルを設定（スクロールモードを垂直に固定）
                                    viewer.style.flexDirection = 'column';
                                    logToVSCode('ビューアのスタイルを設定: 垂直スクロールモード', true);
                                }, 500);
                                
                                // 最初のページがレンダリングされたことを確認してから準備完了を通知
                                const checkFirstPageRendered = (attempts = 0) => {
                                    const firstPageElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                    if (firstPageElement) {
                                        logToVSCode('最初のページのレンダリングを確認しました', true);
                                        
                                        // PDFビューアの準備完了を通知
                                        vscode.postMessage({
                                            command: 'ready',
                                            totalPages: pdf.numPages,
                                            currentPage: pageNum
                                        });
                                        logToVSCode('PDFビューアの準備完了通知を送信しました', true);
                                        
                                        // URLからページ番号を抽出して直接ジャンプ
                                        const urlHash = window.location.hash;
                                        if (urlHash) {
                                            // #page=N 形式のチェック
                                            const pageMatch = urlHash.match(/#page=(\d+)/i);
                                            if (pageMatch && pageMatch[1]) {
                                                const pageFromHash = parseInt(pageMatch[1], 10);
                                                if (pageFromHash > 0 && pageFromHash <= pdf.numPages) {
                                                    logToVSCode('URLハッシュからページ番号 ' + pageFromHash + ' を検出しました。直接ジャンプします。');
                                                    pageNum = pageFromHash;
                                                    pageInput.value = pageNum;
                                                    
                                                    // 少し遅延させてからジャンプ
                                                    setTimeout(() => {
                                                        const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                                        if (targetElement) {
                                                            viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                                            logToVSCode('ページ ' + pageNum + ' に直接スクロールしました（URLハッシュから）');
                                                        }
                                                    }, 500);
                                                }
                                            }
                                            // #N 形式のチェック
                                            else {
                                                const hashMatch = urlHash.match(/#(\d+)$/);
                                                if (hashMatch && hashMatch[1]) {
                                                    const pageFromHash = parseInt(hashMatch[1], 10);
                                                    if (pageFromHash > 0 && pageFromHash <= pdf.numPages) {
                                                        logToVSCode('URLハッシュからページ番号 ' + pageFromHash + ' を検出しました。直接ジャンプします。');
                                                        pageNum = pageFromHash;
                                                        pageInput.value = pageNum;
                                                        
                                                        // 少し遅延させてからジャンプ
                                                        setTimeout(() => {
                                                            const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                                            if (targetElement) {
                                                                viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                                                logToVSCode('ページ ' + pageNum + ' に直接スクロールしました（URLハッシュから）');
                                                            }
                                                        }, 500);
                                                    }
                                                }
                                            }
                                        }
                                    } else if (attempts < 10) {
                                        logToVSCode('最初のページのレンダリングを待機中... (' + (attempts + 1) + '/10)');
                                        setTimeout(() => checkFirstPageRendered(attempts + 1), 300);
                                    } else {
                                        logToVSCode('最初のページのレンダリングタイムアウト。準備完了を通知します。');
                                        
                                        // PDFビューアの準備完了を通知
                                        vscode.postMessage({
                                            command: 'ready',
                                            totalPages: pdf.numPages,
                                            currentPage: pageNum
                                        });
                                        logToVSCode('PDFビューアの準備完了通知を送信しました');
                                    }
                                };
                                
                                // 最初のページのレンダリングを確認
                                setTimeout(() => checkFirstPageRendered(), 500);
                            })
                            .catch(function(error) {
                                loadingMessage.textContent = 'PDFの読み込みに失敗しました: ' + error;
                                reportErrorToVSCode(error);
                                console.error('PDFの読み込みエラー:', error);
                            });

                        // ページをレンダリングするキューに追加
                        function queueRenderPage(num) {
                            if (renderQueue.indexOf(num) === -1) {
                                renderQueue.push(num);
                                if (!isRendering) {
                                    processRenderQueue();
                                }
                            }
                        }

                        // レンダリングキューを処理
                        function processRenderQueue() {
                            if (renderQueue.length === 0) {
                                isRendering = false;
                                return;
                            }
                            
                            isRendering = true;
                            const num = renderQueue.shift();
                            renderPage(num).then(() => {
                                processRenderQueue();
                            }).catch(error => {
                                reportErrorToVSCode(error);
                                console.error('ページレンダリングエラー:', error);
                                processRenderQueue();
                            });
                        }

                        // ページをレンダリング
                        function renderPage(num) {
                            return new Promise((resolve, reject) => {
                                if (pagesCache[num]) {
                                    // すでにレンダリング済みのページがある場合
                                    if (!pagesCache[num].parentNode) {
                                        // ページを正しい位置に挿入する
                                        insertPageInOrder(pagesCache[num], num);
                                    }
                                    resolve();
                                    return;
                                }
                                
                                pdfDoc.getPage(num).then(function(page) {
                                    const viewport = page.getViewport({ scale });
                                    
                                    // ページ用のdiv要素を作成
                                    const pageDiv = document.createElement('div');
                                    pageDiv.className = 'page';
                                    pageDiv.dataset.pageNumber = num;
                                    pageDiv.style.width = viewport.width + 'px';
                                    pageDiv.style.height = viewport.height + 'px';
                                    pageDiv.style.position = 'relative';
                                    
                                    // ページ用のcanvas要素を作成
                                    const canvas = document.createElement('canvas');
                                    const context = canvas.getContext('2d', { alpha: false });
                                    canvas.width = viewport.width;
                                    canvas.height = viewport.height;
                                    pageDiv.appendChild(canvas);
                                    
                                    // ページをレンダリング
                                    const renderContext = {
                                        canvasContext: context,
                                        viewport: viewport
                                    };
                                    
                                    const renderTask = page.render(renderContext);
                                    renderTask.promise.then(function() {
                                        pagesCache[num] = pageDiv;
                                        // ページを正しい位置に挿入する
                                        insertPageInOrder(pageDiv, num);
                                        
                                        // 不要なページをアンロード（メモリ最適化）
                                        optimizeMemory(num);
                                        
                                        resolve();
                                    }).catch(function(error) {
                                        reportErrorToVSCode(error);
                                        console.error('ページレンダリングエラー:', error);
                                        reject(error);
                                    });
                                }).catch(function(error) {
                                    reportErrorToVSCode(error);
                                    console.error('ページ取得エラー:', error);
                                    reject(error);
                                });
                            });
                        }

                        // メモリ使用量を最適化
                        function optimizeMemory(currentNum) {
                            Object.keys(pagesCache).forEach(key => {
                                const keyNum = parseInt(key);
                                if (Math.abs(keyNum - currentNum) > 5) { // 現在のページから5ページ以上離れたページをアンロード
                                    if (pagesCache[key] && pagesCache[key].parentNode) {
                                        pagesCache[key].parentNode.removeChild(pagesCache[key]);
                                    }
                                }
                            });
                        }

                        // 前のページに移動
                        function goPrevPage() {
                            if (pageNum <= 1) return;
                            pageNum--;
                            pageInput.value = pageNum;
                            queueRenderPage(pageNum);
                            
                            // ページのレンダリングが完了した後、そのページにスクロールする
                            // より直接的なスクロール方法を使用
                            const tryScrollToPage = (attempts = 0) => {
                                // 指定されたページの要素を取得
                                const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                
                                if (targetElement) {
                                    // 絶対位置でスクロール
                                    if (targetElement.offsetHeight > viewerContainer.clientHeight) {
                                        // 大きいページは上部に合わせる（少し余白を持たせる）
                                        viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                    } else {
                                        // 小さいページは中央に表示
                                        const centerOffset = (viewerContainer.clientHeight - targetElement.offsetHeight) / 2;
                                        viewerContainer.scrollTop = targetElement.offsetTop - centerOffset;
                                    }
                                    
                                    logToVSCode('ページ ' + pageNum + ' に絶対位置でスクロールしました');
                                } else if (attempts < 5) { // 最大5回試行
                                    logToVSCode('ページ ' + pageNum + ' の要素が見つかりません。再試行します... (' + (attempts + 1) + '/5)');
                                    // ページがまだレンダリングされていない可能性があるため、再試行
                                    setTimeout(() => tryScrollToPage(attempts + 1), 300); // 300ms後に再試行
                                } else {
                                    logToVSCode('ページ ' + pageNum + ' の要素が見つかりませんでした。すべてのページを再レンダリングします。');
                                    // すべてのページを再レンダリング
                                    refreshAllPages();
                                    // 最後にもう一度スクロールを試みる
                                    setTimeout(() => {
                                        const element = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                        if (element) {
                                            if (element.offsetHeight > viewerContainer.clientHeight) {
                                                // 大きいページは上部に合わせる（少し余白を持たせる）
                                                viewerContainer.scrollTop = element.offsetTop - 40;
                                            } else {
                                                // 小さいページは中央に表示
                                                const centerOffset = (viewerContainer.clientHeight - element.offsetHeight) / 2;
                                                viewerContainer.scrollTop = element.offsetTop - centerOffset;
                                            }
                                            
                                            logToVSCode('ページ ' + pageNum + ' に直接スクロールしました（再レンダリング後）');
                                        }
                                    }, 800);
                                }
                            };
                            
                            // 初回試行（500ms待機）
                            setTimeout(() => tryScrollToPage(), 500);
                        }

                        // 次のページに移動
                        function goNextPage() {
                            if (pageNum >= pdfDoc.numPages) return;
                            pageNum++;
                            pageInput.value = pageNum;
                            queueRenderPage(pageNum);
                            
                            // ページのレンダリングが完了した後、そのページにスクロールする
                            // より直接的なスクロール方法を使用
                            const tryScrollToPage = (attempts = 0) => {
                                // 指定されたページの要素を取得
                                const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                
                                if (targetElement) {
                                    // 絶対位置でスクロール
                                    if (targetElement.offsetHeight > viewerContainer.clientHeight) {
                                        // 大きいページは上部に合わせる（少し余白を持たせる）
                                        viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                    } else {
                                        // 小さいページは中央に表示
                                        const centerOffset = (viewerContainer.clientHeight - targetElement.offsetHeight) / 2;
                                        viewerContainer.scrollTop = targetElement.offsetTop - centerOffset;
                                    }
                                    
                                    logToVSCode('ページ ' + pageNum + ' に絶対位置でスクロールしました');
                                } else if (attempts < 5) { // 最大5回試行
                                    logToVSCode('ページ ' + pageNum + ' の要素が見つかりません。再試行します... (' + (attempts + 1) + '/5)');
                                    // ページがまだレンダリングされていない可能性があるため、再試行
                                    setTimeout(() => tryScrollToPage(attempts + 1), 300); // 300ms後に再試行
                                } else {
                                    logToVSCode('ページ ' + pageNum + ' の要素が見つかりませんでした。すべてのページを再レンダリングします。');
                                    // すべてのページを再レンダリング
                                    refreshAllPages();
                                    // 最後にもう一度スクロールを試みる
                                    setTimeout(() => {
                                        const element = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                        if (element) {
                                            if (element.offsetHeight > viewerContainer.clientHeight) {
                                                // 大きいページは上部に合わせる（少し余白を持たせる）
                                                viewerContainer.scrollTop = element.offsetTop - 40;
                                            } else {
                                                // 小さいページは中央に表示
                                                const centerOffset = (viewerContainer.clientHeight - element.offsetHeight) / 2;
                                                viewerContainer.scrollTop = element.offsetTop - centerOffset;
                                            }
                                            
                                            logToVSCode('ページ ' + pageNum + ' に直接スクロールしました（再レンダリング後）');
                                        }
                                    }, 800);
                                }
                            };
                            
                            // 初回試行（500ms待機）
                            setTimeout(() => tryScrollToPage(), 500);
                        }

                        // 拡大
                        function zoomIn() {
                            if (scale >= 3.0) return;
                            scale *= 1.2;
                            resetZoomButtons();
                            refreshAllPages();
                        }

                        // 縮小
                        function zoomOut() {
                            if (scale <= 0.5) return;
                            scale /= 1.2;
                            resetZoomButtons();
                            refreshAllPages();
                        }

                        // ズームをリセット
                        function resetZoom() {
                            scale = 1.5;
                            resetZoomButtons();
                            resetZoomButton.classList.add('active');
                            refreshAllPages();
                        }

                        // 幅に合わせる
                        function fitToWidth() {
                            if (!pdfDoc) return;
                            
                            pdfDoc.getPage(pageNum).then(function(page) {
                                const viewport = page.getViewport({ scale: 1.0 });
                                scale = (containerWidth) / viewport.width;
                                
                                resetZoomButtons();
                                fitWidthButton.classList.add('active');
                                
                                refreshAllPages();
                            });
                        }

                        // ズームボタンのリセット
                        function resetZoomButtons() {
                            fitWidthButton.classList.remove('active');
                            resetZoomButton.classList.remove('active');
                        }

                        // ページを正しい順序で挿入する関数
                        function insertPageInOrder(pageDiv, pageNumber) {
                            // すでに同じページが表示されている場合は何もしない
                            const existingPage = document.querySelector('.page[data-page-number="' + pageNumber + '"]');
                            if (existingPage === pageDiv) {
                                return;
                            }
                            
                            // 既存のページを削除（重複を避けるため）
                            if (existingPage) {
                                existingPage.parentNode.removeChild(existingPage);
                            }
                            
                            // ページを正しい位置に挿入
                            const pages = viewer.querySelectorAll('.page');
                            let inserted = false;
                            
                            for (let i = 0; i < pages.length; i++) {
                                const currentPage = pages[i];
                                const currentPageNum = parseInt(currentPage.dataset.pageNumber, 10);
                                
                                if (currentPageNum > pageNumber) {
                                    // 現在のページの前に挿入
                                    viewer.insertBefore(pageDiv, currentPage);
                                    inserted = true;
                                    break;
                                }
                            }
                            
                            // 最後に挿入（他のすべてのページより大きい番号の場合）
                            if (!inserted) {
                                viewer.appendChild(pageDiv);
                            }
                        }

                        // すべてのページを更新
                        function refreshAllPages() {
                            // キャッシュをクリア
                            pagesCache = {};
                            viewer.innerHTML = '';
                            renderQueue = [];
                            isRendering = false;
                            
                            // 現在のページとその前後のページを優先的にレンダリング
                            const pagesToRender = [];
                            
                            // 現在のページを最初にレンダリング
                            pagesToRender.push(pageNum);
                            
                            // 前後のページを追加
                            for (let i = 1; i <= 2; i++) {
                                if (pageNum + i <= pdfDoc.numPages) {
                                    pagesToRender.push(pageNum + i);
                                }
                                if (pageNum - i >= 1) {
                                    pagesToRender.push(pageNum - i);
                                }
                            }
                            
                            // 残りのページを追加
                            for (let i = 1; i <= pdfDoc.numPages; i++) {
                                if (!pagesToRender.includes(i)) {
                                    pagesToRender.push(i);
                                }
                            }
                            
                            // ページをレンダリングキューに追加
                            pagesToRender.forEach(i => queueRenderPage(i));
                            
                            // 現在のページが表示されるまで待機してからスクロール
                            const scrollToCurrentPage = (attempts = 0) => {
                                // 現在のページの要素を探す
                                const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                
                                if (targetElement) {
                                    // 絶対位置でスクロール
                                    if (targetElement.offsetHeight > viewerContainer.clientHeight) {
                                        // 大きいページは上部に合わせる（少し余白を持たせる）
                                        viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                    } else {
                                        // 小さいページは中央に表示
                                        const centerOffset = (viewerContainer.clientHeight - targetElement.offsetHeight) / 2;
                                        viewerContainer.scrollTop = targetElement.offsetTop - centerOffset;
                                    }
                                    
                                    logToVSCode('ページ ' + pageNum + ' に絶対位置でスクロールしました（再レンダリング後）');
                                } else if (attempts < 10) { // 最大10回試行
                                    // ページがまだレンダリングされていない可能性があるため、再試行
                                    setTimeout(() => scrollToCurrentPage(attempts + 1), 200); // 200ms後に再試行
                                } else {
                                    logToVSCode('ページ ' + pageNum + ' の要素が見つかりませんでした。スクロールに失敗しました。');
                                }
                            };
                            
                            // 初回スクロール試行（500ms待機）
                            setTimeout(() => scrollToCurrentPage(), 500);
                        }

                        // イベントリスナーの設定
                        prevButton.addEventListener('click', goPrevPage);
                        nextButton.addEventListener('click', goNextPage);
                        zoomInButton.addEventListener('click', zoomIn);
                        zoomOutButton.addEventListener('click', zoomOut);
                        resetZoomButton.addEventListener('click', resetZoom);
                        fitWidthButton.addEventListener('click', fitToWidth);
                        
                        pageInput.addEventListener('change', function() {
                            const newPage = parseInt(pageInput.value);
                            if (newPage >= 1 && newPage <= pdfDoc.numPages && newPage !== pageNum) {
                                pageNum = newPage;
                                
                                // 現在のページとその前後のページを優先的にレンダリング
                                queueRenderPage(pageNum);
                                
                                // 前後のページも追加
                                if (pageNum > 1) {
                                    queueRenderPage(pageNum - 1);
                                }
                                if (pageNum < pdfDoc.numPages) {
                                    queueRenderPage(pageNum + 1);
                                }
                                
                                // 指定されたページにスクロール
                                const scrollToPage = (attempts = 0) => {
                                    const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                    
                                    if (targetElement) {
                                        // 絶対位置でスクロール
                                        if (targetElement.offsetHeight > viewerContainer.clientHeight) {
                                            // 大きいページは上部に合わせる（少し余白を持たせる）
                                            viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                        } else {
                                            // 小さいページは中央に表示
                                            const centerOffset = (viewerContainer.clientHeight - targetElement.offsetHeight) / 2;
                                            viewerContainer.scrollTop = targetElement.offsetTop - centerOffset;
                                        }
                                        
                                        logToVSCode('ページ ' + pageNum + ' に絶対位置でスクロールしました');
                                    } else if (attempts < 5) {
                                        setTimeout(() => scrollToPage(attempts + 1), 200);
                                    } else {
                                        logToVSCode('ページ ' + pageNum + ' の要素が見つかりませんでした。すべてのページを再レンダリングします。');
                                        refreshAllPages();
                                        // 最後にもう一度スクロールを試みる
                                        setTimeout(() => {
                                            const element = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                            if (element) {
                                                if (element.offsetHeight > viewerContainer.clientHeight) {
                                                    // 大きいページは上部に合わせる（少し余白を持たせる）
                                                    viewerContainer.scrollTop = element.offsetTop - 40;
                                                } else {
                                                    // 小さいページは中央に表示
                                                    const centerOffset = (viewerContainer.clientHeight - element.offsetHeight) / 2;
                                                    viewerContainer.scrollTop = element.offsetTop - centerOffset;
                                                }
                                                
                                                logToVSCode('ページ ' + pageNum + ' に直接スクロールしました（再レンダリング後）');
                                            }
                                        }, 500);
                                    }
                                };
                                
                                setTimeout(() => scrollToPage(), 300);
                            } else {
                                pageInput.value = pageNum;
                            }
                        });
                        
                        // キーボードショートカット
                        document.addEventListener('keydown', function(e) {
                            if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
                                goPrevPage();
                                e.preventDefault();
                            } else if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
                                goNextPage();
                                e.preventDefault();
                            } else if (e.key === '+' || e.key === '=') {
                                zoomIn();
                                e.preventDefault();
                            } else if (e.key === '-') {
                                zoomOut();
                                e.preventDefault();
                            }
                        });

                        // URLハッシュフラグメントの変更を監視
                        window.addEventListener('hashchange', function() {
                            logToVSCode('URLハッシュが変更されました: ' + window.location.hash);
                            
                            // ハッシュフラグメントからページ番号を抽出
                            const urlHash = window.location.hash;
                            if (urlHash) {
                                // URLをデコード
                                let decodedHash;
                                try {
                                    decodedHash = decodeURIComponent(urlHash);
                                    logToVSCode('デコード後のURLハッシュ: ' + decodedHash);
                                } catch (e) {
                                    logToVSCode('URLハッシュのデコードに失敗: ' + e);
                                    decodedHash = urlHash;
                                }
                                
                                // #page=N 形式のチェック（デコード後）
                                const pageMatch = decodedHash.match(/#page=(\d+)/i);
                                if (pageMatch && pageMatch[1]) {
                                    const newPage = parseInt(pageMatch[1], 10);
                                    logToVSCode('デコードされたURLハッシュから抽出したページ番号(#page=N形式): ' + newPage);
                                    
                                    // ページ番号が有効な範囲内かチェック
                                    if (newPage >= 1 && newPage <= pdfDoc.numPages && newPage !== pageNum) {
                                        pageNum = newPage;
                                        pageInput.value = pageNum;
                                        
                                        // ページをレンダリングしてスクロール
                                        queueRenderPage(pageNum);
                                        
                                        // 前後のページも追加
                                        if (pageNum > 1) {
                                            queueRenderPage(pageNum - 1);
                                        }
                                        if (pageNum < pdfDoc.numPages) {
                                            queueRenderPage(pageNum + 1);
                                        }
                                        
                                        // 指定されたページにスクロール
                                        setTimeout(() => {
                                            const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                            if (targetElement) {
                                                viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                                logToVSCode('ページ ' + pageNum + ' にスクロールしました（ハッシュ変更による）');
                                            }
                                        }, 300);
                                    }
                                }
                                // #N 形式のチェック（デコード後）
                                else {
                                    const hashMatch = decodedHash.match(/#(\d+)$/);
                                    if (hashMatch && hashMatch[1]) {
                                        const newPage = parseInt(hashMatch[1], 10);
                                        logToVSCode('デコードされたURLハッシュから抽出したページ番号(#N形式): ' + newPage);
                                        
                                        // ページ番号が有効な範囲内かチェック
                                        if (newPage >= 1 && newPage <= pdfDoc.numPages && newPage !== pageNum) {
                                            pageNum = newPage;
                                            pageInput.value = pageNum;
                                            
                                            // ページをレンダリングしてスクロール
                                            queueRenderPage(pageNum);
                                            
                                            // 前後のページも追加
                                            if (pageNum > 1) {
                                                queueRenderPage(pageNum - 1);
                                            }
                                            if (pageNum < pdfDoc.numPages) {
                                                queueRenderPage(pageNum + 1);
                                            }
                                            
                                            // 指定されたページにスクロール
                                            setTimeout(() => {
                                                const targetElement = document.querySelector('.page[data-page-number="' + pageNum + '"]');
                                                if (targetElement) {
                                                    viewerContainer.scrollTop = targetElement.offsetTop - 40;
                                                    logToVSCode('ページ ' + pageNum + ' にスクロールしました（ハッシュ変更による）');
                                                }
                                            }, 300);
                                        }
                                    }
                                }
                            }
                        });
                    } catch (error) {
                        reportErrorToVSCode(error);
                        document.getElementById('loadingMessage').textContent = 'PDFの初期化に失敗しました: ' + error;
                        console.error('PDF初期化エラー:', error);
                    }
                </script>
            </body>
            </html>
        `;
    }
}