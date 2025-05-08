import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';

// ハンドラーのインポート
import { PdfHandler } from './handlers/pdfHandler';
import { WordHandler } from './handlers/wordHandler';
import { ExcelHandler } from './handlers/excelHandler';
import { PowerPointHandler } from './handlers/powerPointHandler';
import { ImageHandler } from './handlers/imageHandler';
import { TextHandler } from './handlers/textHandler';
import { CustomHandler } from './handlers/customHandler';

// ユーティリティのインポート
import {
    isPdfFile,
    isWordFile,
    isExcelFile,
    isPowerPointFile,
    isImageFile,
    isTextFile
} from './utils/fileUtils';

/**
 * 拡張機能がアクティブ化されたときに呼び出される
 */
export function activate(context: vscode.ExtensionContext) {
    console.log('ContractMate拡張機能がアクティブ化されました！');

    // 各ハンドラーのインスタンスを作成
    const pdfHandler = new PdfHandler(context);
    const wordHandler = new WordHandler(context);
    const excelHandler = new ExcelHandler(context);
    const powerPointHandler = new PowerPointHandler(context);
    const imageHandler = new ImageHandler(context);
    const textHandler = new TextHandler(context);
    const customHandler = new CustomHandler(context);

    // 各ハンドラーを登録
    const pdfDisposables = pdfHandler.register();
    const wordDisposables = wordHandler.register();
    const excelDisposables = excelHandler.register();
    const powerPointDisposables = powerPointHandler.register();
    const imageDisposables = imageHandler.register();
    const textDisposables = textHandler.register();
    const customDisposables = customHandler.register();

    // Word、Excel、PowerPointファイル用のカスタムエディタプロバイダを登録
    const wordEditorProvider = vscode.window.registerCustomEditorProvider(
        'contractmate.wordViewer',
        new ExternalAppEditorProvider(uri => {
            vscode.commands.executeCommand('contractmate.openWord', uri);
        }),
        { supportsMultipleEditorsPerDocument: false }
    );

    const excelEditorProvider = vscode.window.registerCustomEditorProvider(
        'contractmate.excelViewer',
        new ExternalAppEditorProvider(uri => {
            vscode.commands.executeCommand('contractmate.openExcel', uri);
        }),
        { supportsMultipleEditorsPerDocument: false }
    );

    const powerPointEditorProvider = vscode.window.registerCustomEditorProvider(
        'contractmate.powerPointViewer',
        new ExternalAppEditorProvider(uri => {
            vscode.commands.executeCommand('contractmate.openPowerPoint', uri);
        }),
        { supportsMultipleEditorsPerDocument: false }
    );

    // ファイルを開くコマンドを登録
    const openFileCommand = vscode.commands.registerCommand(
        'contractmate.openFile',
        async (uri: vscode.Uri) => {
            await openFile(uri);
        }
    );
    
    // Markdownファイル内のリンクを処理するコマンドを登録
    const handleMarkdownLinkCommand = vscode.commands.registerCommand(
        'contractmate.handleMarkdownLink',
        async (link: string) => {
            await handleMarkdownLink(link);
        }
    );

    // ファイルエクスプローラーのコンテキストメニューからファイルを開くコマンドを登録
    const openFileFromExplorerCommand = vscode.commands.registerCommand(
        'contractmate.openFileFromExplorer',
        async (uri: vscode.Uri) => {
            await openFile(uri);
        }
    );

    // 設定を開くコマンドを登録
    const openSettingsCommand = vscode.commands.registerCommand(
        'contractmate.openSettings',
        () => {
            vscode.commands.executeCommand('workbench.action.openSettings', 'contractmate');
        }
    );

    // ファイルを開く関数の参照
    const openFileFunc = openFile;

    // ファイルが開かれたときのイベントリスナーを登録
    const fileOpenListener = vscode.workspace.onDidOpenTextDocument(async (document) => {
        // 特定の拡張子のファイルが通常のテキストエディタで開かれた場合、
        // 適切なビューアーで開き直す
        const uri = document.uri;
        
        // ファイルスキームのみを処理（untitledなどは除外）
        if (uri.scheme === 'file') {
            // 1. PDFやJPEGについてはVSCode上でタブで開く
            // 2. 文字データ（md等）についてもタブで開く
            // 3. docxやexcel等の他のファイルについてはOS標準設定のアプリケーションで開く
            if (isPdfFile(uri) || isImageFile(uri)) {
                // PDFや画像ファイルが通常のテキストエディタで開かれた場合、
                // カスタムエディタで開き直す
                const editors = vscode.window.visibleTextEditors.filter(
                    editor => editor.document.uri.toString() === uri.toString()
                );
                
                if (editors.length > 0) {
                    // エディタを閉じる前に適切なビューアーで開く
                    await openFile(uri);
                    
                    // 少し遅延させてからエディタを閉じる（競合を避けるため）
                    setTimeout(() => {
                        vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                    }, 100);
                }
            } else if (isWordFile(uri) || isExcelFile(uri) || isPowerPointFile(uri)) {
                // Word、Excel、PowerPointファイルが通常のテキストエディタで開かれた場合、
                // OS標準設定のアプリケーションで開き直す
                const editors = vscode.window.visibleTextEditors.filter(
                    editor => editor.document.uri.toString() === uri.toString()
                );
                
                if (editors.length > 0) {
                    // エディタを閉じる前に適切なビューアーで開く
                    await openFile(uri);
                    
                    // 少し遅延させてからエディタを閉じる（競合を避けるため）
                    setTimeout(() => {
                        vscode.commands.executeCommand('workbench.action.closeActiveEditor');
                    }, 100);
                }
            }
        }
    });
    
    // Markdownプレビューが開かれたときのイベントリスナーを登録
    // const markdownPreviewListener = vscode.window.registerWebviewPanelSerializer('markdown.preview', {
    //     async deserializeWebviewPanel(webviewPanel: vscode.WebviewPanel, state: any) {
    //         // Markdownプレビューが復元されたときの処理
    //         setupMarkdownLinkInterceptor(webviewPanel);
    //     }
    // });
    
    // Markdownプレビューが作成されたときのイベントリスナーを登録
    // const markdownPreviewProvider = vscode.workspace.registerTextDocumentContentProvider('markdown-preview', {
    //     provideTextDocumentContent(uri: vscode.Uri): string {
    //         return ''; // 実際のコンテンツはVSCodeが提供
    //     }
    // });
    
    // Markdownファイルが開かれたときのイベントを監視
    // const markdownOpenListener = vscode.workspace.onDidOpenTextDocument((document) => {
    //     if (document.languageId === 'markdown') {
    //         // Markdownファイルが開かれたときに、少し遅延させてからプレビューを探す
    //         setTimeout(() => {
    //             // アクティブなWebviewパネルを取得
    //             const activeEditor = vscode.window.activeTextEditor;
    //             if (activeEditor && activeEditor.document.uri.toString() === document.uri.toString()) {
    //                 // コマンドパレットからMarkdownプレビューを開くコマンドを実行
    //                 vscode.commands.executeCommand('markdown.showPreview');
    //             }
    //         }, 500);
    //     }
    // });

    // すべてのDisposableをコンテキストに追加
    context.subscriptions.push(
        ...pdfDisposables,
        ...wordDisposables,
        ...excelDisposables,
        ...powerPointDisposables,
        ...imageDisposables,
        ...textDisposables,
        ...customDisposables,
        openFileCommand,
        openFileFromExplorerCommand,
        openSettingsCommand,
        handleMarkdownLinkCommand,
        fileOpenListener,
        // markdownPreviewListener,
        // markdownPreviewProvider,
        // markdownOpenListener,
        wordEditorProvider,
        excelEditorProvider,
        powerPointEditorProvider
    );

    // 拡張機能の設定を登録
    registerSettings();
}

/**
 * 拡張機能の設定を登録する
 */
function registerSettings(): void {
    // 設定の変更を監視
    vscode.workspace.onDidChangeConfiguration(event => {
        if (event.affectsConfiguration('contractmate')) {
            // 設定が変更された場合の処理
            console.log('ContractMateの設定が変更されました');
        }
    });
}

/**
 * Markdownプレビューパネルにリンクインターセプターを設定する
 */
function setupMarkdownLinkInterceptor(panel: vscode.WebviewPanel): void {
    // WebViewにスクリプトを注入してリンクをインターセプト
    const originalHtml = panel.webview.html;
    
    // リンクインターセプターのスクリプトを追加
    const script = `
        <script>
            (function() {
                // リンククリックイベントをキャプチャ
                document.addEventListener('click', function(event) {
                    let target = event.target;
                    
                    // リンク要素を見つける
                    while (target && target.tagName !== 'A') {
                        target = target.parentElement;
                    }
                    
                    if (target && target.tagName === 'A') {
                        const href = target.getAttribute('href');
                        
                        // すべてのリンクをインターセプト
                        if (href) {
                            // 外部リンク（http://, https://）は除外
                            if (!href.startsWith('http://') && !href.startsWith('https://')) {
                                event.preventDefault();
                                
                                // 絶対パスかどうかを確認（C:\などで始まる場合）
                                const isAbsolutePath = /^[a-zA-Z]:[\\\\/]/.test(href);
                                
                                // VSCodeにメッセージを送信
                                const vscode = acquireVsCodeApi();
                                vscode.postMessage({
                                    command: 'handleLink',
                                    link: href,
                                    isAbsolutePath: isAbsolutePath
                                });
                                
                                console.log('リンクをインターセプトしました:', href, '絶対パス:', isAbsolutePath);
                            }
                        }
                    }
                });
                
                // WebViewからのメッセージを処理
                window.addEventListener('message', function(event) {
                    const message = event.data;
                    if (message.command === 'handleLink') {
                        // VSCodeコマンドを実行
                        vscode.postMessage({
                            command: 'executeCommand',
                            commandId: 'contractmate.handleMarkdownLink',
                            args: [message.link, message.isAbsolutePath]
                        });
                    }
                });
            })();
        </script>
    `;
    
    // HTMLにスクリプトを追加
    if (originalHtml) {
        panel.webview.html = originalHtml.replace('</body>', `${script}</body>`);
    }
    
    // WebViewからのメッセージを処理
    panel.webview.onDidReceiveMessage(message => {
        if (message.command === 'handleLink') {
            console.log('Markdownプレビューからリンクメッセージを受信:', message.link, '絶対パス:', message.isAbsolutePath);
            vscode.commands.executeCommand('contractmate.handleMarkdownLink', message.link, message.isAbsolutePath);
        } else if (message.command === 'executeCommand') {
            vscode.commands.executeCommand(message.commandId, ...message.args);
        }
    });
}

/**
 * ファイルを開く関数
 */
/**
 * ファイルの種類を取得する
 */
function getFileType(uri: vscode.Uri): string {
    if (isPdfFile(uri)) {
        return 'pdf';
    } else if (isImageFile(uri)) {
        return 'image';
    } else if (isTextFile(uri)) {
        return 'text';
    } else if (isWordFile(uri)) {
        return 'word';
    } else if (isExcelFile(uri)) {
        return 'excel';
    } else if (isPowerPointFile(uri)) {
        return 'powerpoint';
    } else {
        return 'other';
    }
}

/**
 * ファイルを開く関数
 * @param uri ファイルのURI
 * @param originalUrl 元のURL文字列（ハッシュフラグメントを含む可能性がある）
 */
async function openFile(uri: vscode.Uri, originalUrl?: string): Promise<void> {
    try {
        // ファイルが存在するか確認
        try {
            fs.accessSync(uri.fsPath, fs.constants.F_OK);
            console.log(`ファイルが存在します: ${uri.fsPath}`);
        } catch (error) {
            console.error(`ファイルが存在しません: ${uri.fsPath}`);
            vscode.window.showErrorMessage(`ファイルが見つかりません: ${uri.fsPath}`);
            return;
        }
        
        // ファイルの種類を取得
        const fileType = getFileType(uri);
        console.log(`ファイルタイプ: ${fileType}`);
        
        // ページ番号の初期値を設定
        let page = 1;
        
        // ファイルの種類に応じて適切なハンドラーを呼び出す
        switch (fileType) {
            case 'pdf':
                // PDFファイルの場合 - VSCode上でタブとして開く
                console.log(`PDFファイルを開きます: ${uri.fsPath}`);
                
                // URIの文字列表現を取得（fsPathはハッシュフラグメントを含まない）
                const uriString = uri.toString();
                
                // 元のファイルパスを取得
                const fsPath = uri.fsPath;
                
                // 元のURL文字列（ハッシュフラグメントを含む可能性がある）
                const fullUrl = originalUrl || uriString;
                
                // URIの詳細情報をログ出力
                console.log(`URI詳細情報:`, {
                    original: uriString,
                    fullUrl: fullUrl,
                    scheme: uri.scheme,
                    authority: uri.authority,
                    path: uri.path,
                    query: uri.query,
                    fragment: uri.fragment,
                    fsPath: uri.fsPath
                });
                
                // ファイルパスからファイル名を抽出
                const fileName = path.basename(fsPath);
                console.log(`ファイル名: ${fileName}`);
                
                // ハッシュフラグメントからページ番号を抽出（VSCodeのURIオブジェクトの仕様により、fragmentプロパティが空の場合がある）
                if (uri.fragment) {
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
                
                // URI.fragmentから抽出できなかった場合、元のURL文字列から検索
                if (page === 1 && fullUrl) {
                    console.log(`元のURL文字列を解析: ${fullUrl}`);
                    
                    // #page=N 形式のチェック
                    const pageMatch = fullUrl.match(/#page=(\d+)/i);
                    if (pageMatch && pageMatch[1]) {
                        page = parseInt(pageMatch[1], 10);
                        console.log(`元のURL文字列から抽出したページ番号(#page=N形式): ${page}`);
                    }
                    // #N 形式のチェック
                    else {
                        const hashMatch = fullUrl.match(/#(\d+)$/);
                        if (hashMatch && hashMatch[1]) {
                            page = parseInt(hashMatch[1], 10);
                            console.log(`元のURL文字列から抽出したページ番号(#N形式): ${page}`);
                        }
                    }
                }
                
                // URI.fragmentから抽出できなかった場合、文字列全体から検索
                if (page === 1) {
                    // #page=N 形式のチェック
                    const pageMatch = uriString.match(/#page=(\d+)/i);
                    if (pageMatch && pageMatch[1]) {
                        page = parseInt(pageMatch[1], 10);
                        console.log(`URI文字列から抽出したページ番号(#page=N形式): ${page}`);
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
                
                // ファイル名に直接ページ番号が含まれている場合（例：file.pdf#page=3）
                if (page === 1 && fileName.includes('#')) {
                    console.log(`ファイル名にハッシュフラグメントが含まれています: ${fileName}`);
                    
                    // #page=N 形式のチェック
                    const pageMatch = fileName.match(/#page=(\d+)/i);
                    if (pageMatch && pageMatch[1]) {
                        page = parseInt(pageMatch[1], 10);
                        console.log(`ファイル名から抽出したページ番号(#page=N形式): ${page}`);
                    }
                    // #N 形式のチェック
                    else {
                        const hashMatch = fileName.match(/#(\d+)$/);
                        if (hashMatch && hashMatch[1]) {
                            page = parseInt(hashMatch[1], 10);
                            console.log(`ファイル名から抽出したページ番号(#N形式): ${page}`);
                        }
                    }
                }
                
                // ページ番号が抽出できた場合は、ページ番号を指定してPDFを開く
                if (page > 1) {
                    console.log(`PDFファイルをページ ${page} で開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openPdfAtPage', uri, page);
                } else {
                    // ページ番号が抽出できなかった場合は、通常通りPDFを開く
                    await vscode.commands.executeCommand('contractmate.openPdf', uri);
                }
                break;
                
            case 'image':
                // 画像ファイルの場合 - VSCode上でタブとして開く
                console.log(`画像ファイルを開きます: ${uri.fsPath}`);
                await vscode.commands.executeCommand('contractmate.openImage', uri);
                break;
                
            case 'text':
                // テキストファイルの場合 - VSCode上でタブとして開く
                console.log(`テキストファイルを開きます: ${uri.fsPath}`);
                
                // ページ番号（行番号）が指定されている場合
                if (page > 1) {
                    console.log(`テキストファイルを行 ${page} で開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openTextAtLine', uri, page);
                } else {
                    await vscode.commands.executeCommand('contractmate.openText', uri);
                }
                break;
                
            case 'word':
                // Wordファイルの場合 - OS標準設定のアプリケーションで開く
                console.log(`Wordファイルを開きます: ${uri.fsPath}`);
                
                // ページ番号が指定されている場合
                if (page > 1) {
                    console.log(`Wordファイルをページ ${page} で開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openWordAtPage', uri, page);
                } else {
                    await vscode.commands.executeCommand('contractmate.openWord', uri);
                }
                break;
                
            case 'excel':
                // Excelファイルの場合 - OS標準設定のアプリケーションで開く
                console.log(`Excelファイルを開きます: ${uri.fsPath}`);
                
                // ページ番号が指定されている場合（シート番号として扱う）
                if (page > 1) {
                    // ページ番号をシート名として使用（Sheet1, Sheet2, ...）
                    const sheet = `Sheet${page}`;
                    console.log(`Excelファイルをシート "${sheet}" で開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openExcelAtSheetCell', uri, sheet);
                } else {
                    await vscode.commands.executeCommand('contractmate.openExcel', uri);
                }
                break;
                
            case 'powerpoint':
                // PowerPointファイルの場合 - OS標準設定のアプリケーションで開く
                console.log(`PowerPointファイルを開きます: ${uri.fsPath}`);
                
                // スライド番号が指定されている場合
                if (page > 1) {
                    console.log(`PowerPointファイルをスライド ${page} で開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openPowerPointAtSlide', uri, page);
                } else {
                    await vscode.commands.executeCommand('contractmate.openPowerPoint', uri);
                }
                break;
                
            default:
                // その他のファイルの場合 - OS標準設定のアプリケーションで開く
                console.log(`その他のファイルを開きます: ${uri.fsPath}`);
                await vscode.commands.executeCommand('contractmate.openCustom', uri);
                break;
        }
    } catch (error) {
        console.error(`ファイルを開く処理でエラーが発生しました: ${error}`);
        vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error}`);
    }
}

/**
 * Markdownファイル内のリンクを処理する
 */
async function handleMarkdownLink(link: string, isAbsolutePath?: boolean): Promise<void> {
    try {
        console.log(`Markdownリンクを処理: ${link}, 絶対パス: ${isAbsolutePath}`);
        
        // デバッグ用：リンクの詳細情報を出力
        console.log(`リンクの詳細情報:`);
        console.log(`- 長さ: ${link.length}`);
        console.log(`- 文字コード:`, Array.from(link).map(c => c.charCodeAt(0).toString(16)).join(' '));
        
        // リンクからパスとページ番号を抽出
        let filePath = link;
        let page = 1;
        
        // ?page=N 形式のページ指定があるか確認（クエリパラメータ）
        if (link.includes('?page=')) {
            const parts = link.split('?page=');
            filePath = parts[0];
            page = parseInt(parts[1], 10) || 1;
            console.log(`?page= 形式のページ指定を検出: ページ ${page}`);
        }
        // #page=N 形式のページ指定があるか確認（フラグメント識別子）
        else if (link.includes('#page=')) {
            const parts = link.split('#page=');
            filePath = parts[0];
            page = parseInt(parts[1], 10) || 1;
            console.log(`#page= 形式のページ指定を検出: ページ ${page}`);
            
            // #page=N を ?page=N に変換（VSCodeのURIオブジェクトがハッシュフラグメントを正しく保持しないため）
            filePath = filePath + '?page=' + page;
            console.log(`#page= 形式を ?page= 形式に変換: ${filePath}`);
        } else if (link.includes('#')) {
            // 他の形式のフラグメント識別子がある場合
            const parts = link.split('#');
            filePath = parts[0];
            // フラグメント識別子が数値の場合はページ番号として扱う
            const pageNum = parseInt(parts[1], 10);
            if (!isNaN(pageNum)) {
                page = pageNum;
                console.log(`# 形式のページ指定を検出: ページ ${page}`);
                
                // #N を ?page=N に変換（VSCodeのURIオブジェクトがハッシュフラグメントを正しく保持しないため）
                filePath = filePath + '?page=' + page;
                console.log(`# 形式を ?page= 形式に変換: ${filePath}`);
            }
        }
        
        // パスの正規化と解決
        let uri: vscode.Uri;
        
        // 1. file:/// 形式のURLを処理
        if (filePath.startsWith('file:///')) {
            console.log(`file:/// 形式のURLを検出: ${filePath}`);
            
            try {
                // file:/// プレフィックスを削除してデコード
                const decodedPath = decodeURIComponent(filePath.substring(8)); // 'file:///' の長さは8
                console.log(`デコードされたパス: ${decodedPath}`);
                
                // Windowsパスの修正（先頭の / を削除）
                const fixedPath = decodedPath.startsWith('/') ? decodedPath.substring(1) : decodedPath;
                console.log(`修正されたパス: ${fixedPath}`);
                
                // パスを正規化
                const normalizedPath = path.normalize(fixedPath);
                console.log(`正規化されたパス: ${normalizedPath}`);
                
                // URIを直接作成
                uri = vscode.Uri.file(normalizedPath);
                console.log(`作成されたURI: ${uri.toString()}`);
            } catch (error) {
                console.error(`URLのデコード中にエラーが発生しました: ${error}`);
                throw new Error(`URLのデコード中にエラーが発生しました: ${error}`);
            }
        }
        // 2. Windowsの絶対パス（C:\など）を処理
        else if (filePath.match(/^[a-zA-Z]:\\/) || isAbsolutePath === true) {
            console.log(`Windowsの絶対パスを検出: ${filePath}`);
            
            // パスを正規化
            const normalizedPath = path.normalize(filePath);
            console.log(`正規化されたパス: ${normalizedPath}`);
            
            // 絶対パスの場合は直接URIを作成（ワークスペースパスを追加しない）
            uri = vscode.Uri.file(normalizedPath);
            console.log(`作成されたURI: ${uri.toString()}`);
            
            // URIのfsPathを確認（二重パスになっていないか）
            console.log(`URI.fsPath: ${uri.fsPath}`);
            
            // 二重パスのチェック（C:\が2回出現していないか）
            if (uri.fsPath.match(/[a-zA-Z]:\\.*[a-zA-Z]:\\/)) {
                console.error(`二重パスが検出されました: ${uri.fsPath}`);
                
                // 最後のドライブレター以降の部分を抽出
                const match = uri.fsPath.match(/.*([a-zA-Z]:\\.*)/);
                if (match && match[1]) {
                    const correctedPath = match[1];
                    console.log(`修正されたパス: ${correctedPath}`);
                    uri = vscode.Uri.file(correctedPath);
                    console.log(`修正されたURI: ${uri.toString()}`);
                }
            }
        }
        // 3. 相対パスを処理
        else {
            console.log(`相対パスを処理: ${filePath}`);
            
            // 現在のワークスペースフォルダを基準にする
            const workspaceFolders = vscode.workspace.workspaceFolders;
            if (!workspaceFolders || workspaceFolders.length === 0) {
                throw new Error('ワークスペースフォルダが見つかりません');
            }
            
            const workspacePath = workspaceFolders[0].uri.fsPath;
            console.log(`ワークスペースパス: ${workspacePath}`);
            
            // パスを解決して正規化
            const resolvedPath = path.resolve(workspacePath, filePath);
            const normalizedPath = path.normalize(resolvedPath);
            console.log(`解決されたパス: ${resolvedPath}`);
            console.log(`正規化されたパス: ${normalizedPath}`);
            
            // URIを作成
            uri = vscode.Uri.file(normalizedPath);
            console.log(`作成されたURI: ${uri.toString()}`);
        }
        
        // ファイルが存在するか確認
        // 1. fs.accessSyncを使用
        let fileExists = false;
        try {
            fs.accessSync(uri.fsPath, fs.constants.F_OK);
            console.log(`fs.accessSync: ファイルが存在します: ${uri.fsPath}`);
            fileExists = true;
        } catch (error) {
            console.error(`fs.accessSync: ファイルが存在しません: ${uri.fsPath}`);
            // エラーメッセージはまだ表示しない
        }
        
        // 2. fs.existsSyncを使用（代替方法）
        if (!fileExists) {
            try {
                fileExists = fs.existsSync(uri.fsPath);
                console.log(`fs.existsSync: ファイルが存在${fileExists ? 'します' : 'しません'}: ${uri.fsPath}`);
            } catch (error) {
                console.error(`fs.existsSync: エラーが発生しました: ${error}`);
            }
        }
        
        // 3. 最後の手段：直接ファイルを開いてみる
        if (!fileExists) {
            try {
                // ファイルを読み取りモードで開いてみる（存在確認のみ）
                const fd = fs.openSync(uri.fsPath, 'r');
                fs.closeSync(fd);
                console.log(`fs.openSync: ファイルが存在します: ${uri.fsPath}`);
                fileExists = true;
            } catch (error) {
                console.error(`fs.openSync: ファイルが存在しません: ${uri.fsPath}`);
                vscode.window.showErrorMessage(`ファイルが見つかりません: ${uri.fsPath}`);
                return;
            }
        }
        
        // ファイルの種類を判定
        const fileType = getFileType(uri);
        console.log(`ファイルタイプ: ${fileType}`);
        
        // ファイルの種類に応じて適切なコマンドを実行
        try {
            switch (fileType) {
                case 'pdf':
                    // PDFファイルの場合、ページ番号を指定して開く
                    console.log(`PDFファイルを開きます（ページ ${page}）: ${uri.fsPath}`);
                    
                    // URIオブジェクトにハッシュフラグメントを追加（VSCodeのURIオブジェクトの仕様上、これは機能しない可能性がある）
                    let pdfUri = uri;
                    
                    // ファイル名にハッシュフラグメントを追加（これは表示用のみで、実際のファイルパスには影響しない）
                    if (page > 1) {
                        console.log(`ページ番号 ${page} をURIに追加します`);
                        
                        // ファイル名を取得
                        const fileName = path.basename(uri.fsPath);
                        console.log(`元のファイル名: ${fileName}`);
                        
                        // ハッシュフラグメントを含むファイル名を作成
                        const newFileName = fileName + `#page=${page}`;
                        console.log(`新しいファイル名: ${newFileName}`);
                        
                        // デバッグ情報を出力
                        console.log(`ページ番号 ${page} を指定してPDFを開きます: ${uri.fsPath}`);
                    }
                    
                    // ページ番号を指定してPDFを開く
                    // 元のリンク文字列も渡す
                    // 注意: VSCodeのURIオブジェクトがハッシュフラグメントを正しく保持しないため、
                    // #page=N 形式は ?page=N 形式に変換済み
                    await vscode.commands.executeCommand('contractmate.openPdfAtPage', pdfUri, page, link);
                    break;
                    
                case 'image':
                    // 画像ファイルの場合
                    console.log(`画像ファイルを開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openImage', uri);
                    break;
                    
                case 'text':
                    // テキストファイルの場合
                    console.log(`テキストファイルを開きます: ${uri.fsPath}`);
                    
                    // ページ番号（行番号）が指定されている場合
                    if (page > 1) {
                        console.log(`テキストファイルを行 ${page} で開きます: ${uri.fsPath}`);
                        await vscode.commands.executeCommand('contractmate.openTextAtLine', uri, page);
                    } else {
                        await vscode.commands.executeCommand('contractmate.openText', uri);
                    }
                    break;
                    
                case 'word':
                    // Wordファイルの場合
                    console.log(`Wordファイルを開きます: ${uri.fsPath}`);
                    
                    // ページ番号が指定されている場合
                    if (page > 1) {
                        console.log(`Wordファイルをページ ${page} で開きます: ${uri.fsPath}`);
                        await vscode.commands.executeCommand('contractmate.openWordAtPage', uri, page);
                    }
                    // ブックマークが指定されている場合
                    else if (link.includes('#bookmark=') || link.includes('?bookmark=')) {
                        // ブックマーク名を抽出
                        const bookmarkMatch = link.match(/#bookmark=([^&]+)/i) || link.match(/\?bookmark=([^&]+)/i);
                        if (bookmarkMatch && bookmarkMatch[1]) {
                            const bookmark = bookmarkMatch[1];
                            console.log(`Wordファイルをブックマーク "${bookmark}" で開きます: ${uri.fsPath}`);
                            await vscode.commands.executeCommand('contractmate.openWordAtBookmark', uri, bookmark);
                        } else {
                            await vscode.commands.executeCommand('contractmate.openWord', uri);
                        }
                    } else {
                        await vscode.commands.executeCommand('contractmate.openWord', uri);
                    }
                    break;
                    
                case 'excel':
                    // Excelファイルの場合
                    console.log(`Excelファイルを開きます: ${uri.fsPath}`);
                    
                    // シート名が指定されている場合
                    if (link.includes('#sheet=') || link.includes('?sheet=')) {
                        // シート名を抽出
                        const sheetMatch = link.match(/#sheet=([^!&]+)/i) || link.match(/\?sheet=([^!&]+)/i);
                        let sheet = '';
                        let cell = '';
                        
                        if (sheetMatch && sheetMatch[1]) {
                            sheet = sheetMatch[1];
                            console.log(`シート名を抽出しました: ${sheet}`);
                            
                            // セル位置も抽出
                            const cellMatch = link.match(/#cell=([A-Z0-9]+)/i) || link.match(/[?&]cell=([A-Z0-9]+)/i);
                            if (cellMatch && cellMatch[1]) {
                                cell = cellMatch[1];
                                console.log(`セル位置を抽出しました: ${cell}`);
                            }
                            
                            // シート名!セル位置 形式のチェック
                            const sheetCellMatch = link.match(/#sheet=([^!]+)!([A-Z0-9]+)/i) || link.match(/\?sheet=([^!]+)!([A-Z0-9]+)/i);
                            if (sheetCellMatch && sheetCellMatch[1] && sheetCellMatch[2]) {
                                sheet = sheetCellMatch[1];
                                cell = sheetCellMatch[2];
                                console.log(`シート名とセル位置を抽出しました: シート=${sheet}, セル=${cell}`);
                            }
                            
                            console.log(`Excelファイルをシート "${sheet}"${cell ? ` のセル "${cell}"` : ''} で開きます: ${uri.fsPath}`);
                            await vscode.commands.executeCommand('contractmate.openExcelAtSheetCell', uri, sheet, cell);
                        } else {
                            await vscode.commands.executeCommand('contractmate.openExcel', uri);
                        }
                    }
                    // ページ番号が指定されている場合（シート番号として扱う）
                    else if (page > 1) {
                        // ページ番号をシート名として使用（Sheet1, Sheet2, ...）
                        const sheet = `Sheet${page}`;
                        console.log(`Excelファイルをシート "${sheet}" で開きます: ${uri.fsPath}`);
                        await vscode.commands.executeCommand('contractmate.openExcelAtSheetCell', uri, sheet);
                    } else {
                        await vscode.commands.executeCommand('contractmate.openExcel', uri);
                    }
                    break;
                    
                case 'powerpoint':
                    // PowerPointファイルの場合
                    console.log(`PowerPointファイルを開きます: ${uri.fsPath}`);
                    
                    // スライド番号が指定されている場合
                    if (page > 1) {
                        console.log(`PowerPointファイルをスライド ${page} で開きます: ${uri.fsPath}`);
                        await vscode.commands.executeCommand('contractmate.openPowerPointAtSlide', uri, page);
                    } else {
                        await vscode.commands.executeCommand('contractmate.openPowerPoint', uri);
                    }
                    break;
                    
                default:
                    // その他のファイルの場合
                    console.log(`その他のファイルを開きます: ${uri.fsPath}`);
                    await vscode.commands.executeCommand('contractmate.openCustom', uri);
                    break;
            }
        } catch (error) {
            console.error(`ファイルを開く処理でエラーが発生しました: ${error}`);
            vscode.window.showErrorMessage(`ファイルを開けませんでした: ${error}`);
        }
    } catch (error) {
        vscode.window.showErrorMessage(`リンクを処理できませんでした: ${error}`);
    }
}

/**
 * 拡張機能が非アクティブ化されたときに呼び出される
 */
export function deactivate() {
    console.log('ContractMate拡張機能が非アクティブ化されました');
}

/**
 * 外部アプリケーションでファイルを開くカスタムエディタプロバイダ
 */
class ExternalAppEditorProvider implements vscode.CustomReadonlyEditorProvider {
    constructor(private readonly openHandler: (uri: vscode.Uri) => void) {}

    async openCustomDocument(
        uri: vscode.Uri,
        _openContext: vscode.CustomDocumentOpenContext,
        _token: vscode.CancellationToken
    ): Promise<vscode.CustomDocument> {
        // 外部アプリケーションでファイルを開く
        this.openHandler(uri);
        
        // VSCodeのエディタを閉じる
        setTimeout(() => {
            vscode.commands.executeCommand('workbench.action.closeActiveEditor');
        }, 100);
        
        return { uri, dispose: () => {} };
    }

    async resolveCustomEditor(
        document: vscode.CustomDocument,
        webviewPanel: vscode.WebviewPanel,
        _token: vscode.CancellationToken
    ): Promise<void> {
        // 外部アプリケーションで開くため、WebViewの内容は最小限にする
        webviewPanel.webview.html = `
            <!DOCTYPE html>
            <html lang="ja">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>外部アプリケーションで開いています</title>
                <style>
                    body {
                        display: flex;
                        flex-direction: column;
                        align-items: center;
                        justify-content: center;
                        height: 100vh;
                        font-family: var(--vscode-font-family);
                        color: var(--vscode-editor-foreground);
                        background-color: var(--vscode-editor-background);
                        text-align: center;
                        padding: 20px;
                    }
                    .message {
                        margin-bottom: 20px;
                    }
                </style>
            </head>
            <body>
                <div class="message">
                    <h2>外部アプリケーションでファイルを開いています</h2>
                    <p>このファイルは外部アプリケーションで開かれています。</p>
                    <p>ファイル: ${document.uri.fsPath}</p>
                </div>
            </body>
            </html>
        `;
    }
}
