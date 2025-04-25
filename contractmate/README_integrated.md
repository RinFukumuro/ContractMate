# ContractMate6

ContractMate6は、法律文書の作成と管理を支援するVSCode拡張機能です。PDFビューア、Word文書ナビゲーション、その他のOfficeファイル連携機能を提供します。

## 📋 目次

- [主な機能](#主な機能)
- [PDFビューア](#pdfビューア)
- [テキストファイルナビゲーション](#テキストファイルナビゲーション)
- [Word文書ナビゲーション](#word文書ナビゲーション)
- [Excelファイル連携](#excelファイル連携)
- [PowerPointファイル連携](#powerPointファイル連携)
- [使用方法](#使用方法)
- [設定方法](#設定方法)
- [要件](#要件)
- [既知の問題と対処法](#既知の問題と対処法)
- [リリースノート](#リリースノート)
- [今後の予定](#今後の予定)
- [開発者向け情報](#開発者向け情報)

## 主な機能

ContractMate6は、様々なファイルタイプに対応したページジャンプ機能を提供します。これにより、Markdownファイル内のリンクから特定のページや位置に直接ジャンプできます。
### PDFビューア

ContractMate6は、VSCode内で直接PDFファイルを表示するカスタムPDFビューアを提供します。

- **PDFファイルの表示**: VSCode内でPDFファイルを直接開いて表示
- **ページジャンプ機能**: 特定のページに直接ジャンプ可能
- **ズーム機能**: PDFの拡大・縮小、幅に合わせる機能
- **ページナビゲーション**: 前後のページへの移動、ページ番号指定による移動
- **検索機能**: PDF内のテキスト検索

#### ページジャンプ機能

PDFファイルを開く際に特定のページに直接ジャンプするには、以下の形式でリンクを作成します：

```
file:///path/to/document.pdf#page=2
```

以下の形式もサポートされています：

```
file:///path/to/document.pdf?page=2
file:///path/to/document.pdf#2
```

##### 使用例

Markdownファイル内でPDFファイルの特定のページにリンクする例：

```markdown
[契約書の第2条](file:///C:/Documents/contract.pdf#page=5)
[請求書](file:///C:/Documents/invoice.pdf?page=2)
[報告書の概要](file:///C:/Documents/report.pdf#3)
```

### テキストファイルナビゲーション

ContractMate6は、テキストファイルの表示と行番号指定機能を提供します。

- **テキストファイルの表示**: VSCode内でテキストファイルを表示
- **行番号ジャンプ機能**: 特定の行番号に直接ジャンプ可能
- **シンタックスハイライト**: ファイル拡張子に基づいた構文強調表示

#### 行番号ジャンプ機能

テキストファイルを開く際に特定の行に直接ジャンプするには、以下の形式でリンクを作成します：

```
file:///path/to/document.txt#line=10
```

以下の形式もサポートされています：

```
file:///path/to/document.txt?line=10
file:///path/to/document.txt#10
```

##### 使用例

Markdownファイル内でテキストファイルの特定の行にリンクする例：

```markdown
[設定ファイルの重要な部分](file:///C:/Projects/config.json#line=25)
[エラーが発生している行](file:///C:/Projects/script.js?line=42)
[ログファイルの警告メッセージ](file:///C:/Logs/app.log#100)
```

#### サポートされているテキストファイル形式

- テキストファイル (.txt)
- Markdown (.md)
- JSON (.json)
- XML (.xml)
- HTML (.html)
- CSS (.css)
- JavaScript (.js)
- TypeScript (.ts)
- その他のテキストベースのファイル形式
## Word文書ナビゲーション

ContractMate6は、Word文書の表示と操作を支援する機能を提供します。

### 📝 概要

**Word文書ナビゲーション**機能は、弁護士や法務担当者が日々扱う大量の契約書や法律文書をより効率的に操作するためのツールです。

**こんな悩みはありませんか？**
- 長い契約書の特定条項をクライアントの前ですぐに見つけられない
- 複数の文書を行き来する際に、目的の場所を探すのに時間がかかる
- 会議中に「第7条を見せて」と言われても、すぐに該当箇所を表示できない

このツールを使えば、**ワンコマンドで特定の条項や段落に即座にジャンプ**できます！

### 🚀 主な機能

- **Word文書を自動で開く** - ファイルパスを指定するだけで、Wordが起動して文書が開きます
- **文書構造を分析** - 文書内のすべての段落、見出し、条項を自動的に識別します
- **特定の段落に即座にジャンプ** - 段落番号を指定するだけで、その場所に瞬時に移動します
- **文書情報の表示** - 段落のスタイル、アウトラインレベル、内容を一覧表示します
- **ページジャンプ機能**: 特定のページに直接ジャンプ可能（Windows環境のみ）
- **ブックマークジャンプ機能**: 特定のブックマークに直接ジャンプ可能（Windows環境のみ）

### 基本的な使い方

#### VSCode内からのリンク

Word文書を開く際に特定のページに直接ジャンプするには、以下の形式でリンクを作成します：

```
file:///path/to/document.docx#page=5
```

#### ブックマークジャンプ機能

Word文書を開く際に特定のブックマークに直接ジャンプするには、以下の形式でリンクを作成します：

```
file:///path/to/document.docx#bookmark=BookmarkName
```

##### 使用例

Markdownファイル内でWord文書の特定のページやブックマークにリンクする例：

```markdown
[契約書の第3条](file:///C:/Documents/contract.docx#page=7)
[重要な定義](file:///C:/Documents/agreement.docx#bookmark=Definitions)
```

#### コマンドラインからの使用

より高度な文書ナビゲーション機能を使用するには、コマンドラインから以下のコマンドを実行します：

```
cscript simple_navigator.js "C:\Users\あなたの名前\Documents\契約書.docx" 1
```

これにより、指定した文書が開き、1段落目に移動します。

別の段落に移動したい場合は、コマンドの最後の数字を変更するだけです：

```
cscript simple_navigator.js "C:\Users\あなたの名前\Documents\契約書.docx" 5
```

これで5段落目に移動します。

### 段落番号の見つけ方

段落番号は1から始まり、文書内の各段落（改行で区切られたテキスト）ごとに1つずつ割り当てられます。

**段落番号を確認する方法：**

1. まず、段落番号を知らない場合は、段落番号1を指定して実行します：
   ```
   cscript simple_navigator.js "C:\Users\あなたの名前\Documents\契約書.docx" 1
   ```

2. コマンドプロンプトに文書内のすべての段落情報が表示されます
3. 表示された情報から、移動したい段落の番号を確認できます

例えば、「第3条（契約期間）」という見出しが13番目の段落だと分かれば、次回からは：

```
cscript simple_navigator.js "C:\Users\あなたの名前\Documents\契約書.docx" 13
```

と実行するだけで、即座に第3条に移動できます！

### サポートされているWord文書形式

- Word文書 (.doc, .docx)

### 活用事例

#### 事例1: 契約書レビューの効率化

**Before:** 100ページの契約書から特定の条項を探すのに平均2分かかっていた

**After:** ワンコマンドで即座に目的の条項にジャンプ。時間削減率95%！

#### 事例2: クライアントとの会議での活用

**Before:** 「第8条を見せてください」と言われて慌てて探す場面があった

**After:** 事前に段落番号をメモしておき、リクエストされた瞬間に該当条項を表示できるようになった

#### 事例3: 複数文書の比較作業

**Before:** 複数の契約書の同じ条項を比較するのに、文書を行き来する必要があった

**After:** 複数のコマンドプロンプトを開き、各文書の同じ条項を並べて表示できるようになった
### Excelファイル連携

ContractMate6は、Excelファイルの表示とシート・セル指定機能を提供します。

- **Excelファイルの表示**: OS標準のアプリケーションでExcelファイルを開く
- **シートジャンプ機能**: 特定のシートに直接ジャンプ可能（Windows環境のみ）
- **セルジャンプ機能**: 特定のセルに直接ジャンプ可能（Windows環境のみ）
- **データの可視化**: 表やグラフの表示（将来のバージョンで実装予定）

#### シート・セルジャンプ機能

Excelファイルを開く際に特定のシートとセルに直接ジャンプするには、以下の形式でリンクを作成します：

```
file:///path/to/document.xlsx#sheet=Sheet1!A1
```

シートのみを指定する場合：

```
file:///path/to/document.xlsx#sheet=Sheet1
```

ページ番号をシート番号として使用することもできます：

```
file:///path/to/document.xlsx#page=2
```

##### 使用例

Markdownファイル内でExcelファイルの特定のシートやセルにリンクする例：

```markdown
[財務データ](file:///C:/Documents/finance.xlsx#sheet=Q2_2023!B15)
[売上シート](file:///C:/Documents/sales.xlsx#sheet=Monthly)
[3番目のシート](file:///C:/Documents/data.xlsx#page=3)
```

#### サポートされているExcelファイル形式

- Excelファイル (.xls, .xlsx)

### PowerPointファイル連携

ContractMate6は、PowerPointファイルの表示とスライド指定機能を提供します。

- **PowerPointファイルの表示**: OS標準のアプリケーションでPowerPointファイルを開く
- **スライドジャンプ機能**: 特定のスライドに直接ジャンプ可能（Windows環境のみ）
- **プレゼンテーションモード**: 自動的にスライドショーを開始（将来のバージョンで実装予定）

#### スライドジャンプ機能

PowerPointファイルを開く際に特定のスライドに直接ジャンプするには、以下の形式でリンクを作成します：

```
file:///path/to/document.pptx#slide=3
```

以下の形式もサポートされています：

```
file:///path/to/document.pptx#page=3
```

##### 使用例

Markdownファイル内でPowerPointファイルの特定のスライドにリンクする例：

```markdown
[プレゼンテーションの重要なスライド](file:///C:/Documents/presentation.pptx#slide=5)
[概要スライド](file:///C:/Documents/overview.pptx#page=2)
```

#### サポートされているPowerPointファイル形式

- PowerPointファイル (.ppt, .pptx, .pps, .ppsx)
## 使用方法

### PDFビューアの使用

1. VSCodeでPDFファイルを開きます
2. PDFファイルが自動的にContractMate6のPDFビューアで表示されます
3. ツールバーのボタンを使用して、ズームの調整やページの移動を行います

### ページジャンプ機能の使用方法

Markdown文書内でファイルへのリンクを作成する際に、以下のようにページ番号やその他の位置情報を指定します。これらのリンクをクリックすると、対応するファイルが開き、指定した位置に自動的にジャンプします。

#### 基本的な使用方法

1. Markdownファイル内でリンクを作成します
2. リンク先のファイルパスを指定します
3. ファイルパスの後に`#`または`?`を付けて、ページ番号や位置情報を指定します
4. リンクをクリックすると、ファイルが開き、指定した位置にジャンプします

#### ファイルタイプ別のリンク形式

| ファイルタイプ | リンク形式 | 例 |
|------------|---------|-----|
| PDF | `file:///path/to/file.pdf#page=N`<br>`file:///path/to/file.pdf?page=N`<br>`file:///path/to/file.pdf#N` | `[契約書](file:///C:/Documents/contract.pdf#page=5)` |
| テキスト | `file:///path/to/file.txt#line=N`<br>`file:///path/to/file.txt?line=N`<br>`file:///path/to/file.txt#N` | `[設定](file:///C:/Projects/config.json#line=25)` |
| Word | `file:///path/to/file.docx#page=N`<br>`file:///path/to/file.docx#bookmark=Name` | `[報告書](file:///C:/Documents/report.docx#page=7)` |
| Excel | `file:///path/to/file.xlsx#sheet=Name!Cell`<br>`file:///path/to/file.xlsx#sheet=Name`<br>`file:///path/to/file.xlsx#page=N` | `[データ](file:///C:/Documents/data.xlsx#sheet=Sheet1!A1)` |
| PowerPoint | `file:///path/to/file.pptx#slide=N`<br>`file:///path/to/file.pptx#page=N` | `[プレゼン](file:///C:/Documents/presentation.pptx#slide=3)` |

#### 実際の使用例

以下は、様々なファイルタイプへのリンクを含むMarkdownファイルの例です：

```markdown
# プロジェクト資料

## 契約書

- [契約書本文](file:///C:/Documents/contract.pdf#page=1)
- [第2条（定義）](file:///C:/Documents/contract.pdf#page=3)
- [第10条（秘密保持）](file:///C:/Documents/contract.pdf#page=8)
- [別紙A](file:///C:/Documents/contract.pdf#page=15)

## 技術資料

- [システム設計書](file:///C:/Documents/design.docx#bookmark=Architecture)
- [API仕様書](file:///C:/Projects/api-spec.md#line=50)
- [設定ファイル](file:///C:/Projects/config.json#line=25)

## データ

- [売上データ](file:///C:/Documents/sales.xlsx#sheet=2023Q2!B10)
- [顧客リスト](file:///C:/Documents/customers.xlsx#sheet=Active)

## プレゼンテーション

- [プロジェクト概要](file:///C:/Documents/project.pptx#slide=1)
- [スケジュール](file:///C:/Documents/project.pptx#slide=5)
- [予算](file:///C:/Documents/project.pptx#slide=8)
```

### Word文書ナビゲーションの高度な使用方法

#### コマンドラインツールのファイル

Word文書ナビゲーション機能には以下のファイルが含まれています：

- **simple_navigator.js** - 最も簡単に使えるナビゲーションツール（初心者向け）
- **wordNavigator.js** - より多機能なナビゲーションツール（上級者向け）
- **createSampleDoc.js** - テスト用のサンプル契約書を作成するツール
- **testWordNavigator.js** - 自動テスト用のスクリプト

初めて使う方は **simple_navigator.js** から始めることをお勧めします。

#### テスト用サンプル文書の作成

テスト用のサンプル契約書を作成するには、以下のコマンドを実行します：

```
cscript createSampleDoc.js
```

このコマンドは、テスト用のサンプル契約書（サンプル契約書.docx）を現在のディレクトリに作成します。

#### 自動テスト

Word文書ナビゲーション機能を自動的にテストするには、以下のコマンドを実行します：

```
cscript testWordNavigator.js
```

このコマンドは、サンプル契約書が存在しない場合は自動生成し、WordNavigatorを使用してサンプル契約書を開き、文書の段落構造を表示し、デフォルトで1段落目に移動します。

## 設定方法

ContractMate6は、様々な設定オプションを提供しています。これらの設定は、VSCodeの設定画面から変更できます。

### 設定へのアクセス方法

1. VSCodeのメニューから「ファイル」→「ユーザー設定」→「設定」を選択します
2. 検索ボックスに「contractmate」と入力します
3. ContractMate6の設定が表示されます

または、コマンドパレット（Ctrl+Shift+P）から「ContractMate: 設定を開く」を実行することもできます。

### 主な設定項目

| 設定項目 | 説明 | デフォルト値 |
|--------|------|------------|
| `contractmate.wordAppPath` | Microsoft Wordのパス | 自動検出 |
| `contractmate.excelAppPath` | Microsoft Excelのパス | 自動検出 |
| `contractmate.powerPointAppPath` | Microsoft PowerPointのパス | 自動検出 |
| `contractmate.pdfViewer.defaultZoom` | PDFビューアのデフォルトズーム | 1.0 |
| `contractmate.pdfViewer.scrollMode` | PDFビューアのスクロールモード（0: 垂直, 1: 水平, 2: 折り返し） | 0 |

### Office アプリケーションのパス設定

Word、Excel、PowerPointのページジャンプ機能を使用するには、各アプリケーションのパスを設定する必要があります。Windows環境では、通常以下のようなパスになります：

- Word: `C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE`
- Excel: `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE`
- PowerPoint: `C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE`

これらのパスは、Office のバージョンやインストール場所によって異なる場合があります。

#### 設定例

```json
{
  "contractmate.wordAppPath": "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
  "contractmate.excelAppPath": "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
  "contractmate.powerPointAppPath": "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
}
```

## 要件

- Visual Studio Code 1.60.0以降
- PDF.js（拡張機能に同梱）
- Microsoft Office（Word、Excel、PowerPointのページジャンプ機能を使用する場合）
- Windows環境（Word、Excel、PowerPointの高度なナビゲーション機能を使用する場合）
- Windows Script Host (WSH)（Word文書ナビゲーションのコマンドラインツールを使用する場合）

## 既知の問題と対処法

### 一般的な問題

- **非常に大きなPDFファイルの場合、表示に時間がかかることがあります。**
  - 対処法: PDFファイルを分割するか、より小さなファイルに変換することを検討してください。

- **Word、Excel、PowerPointのページジャンプ機能は、Windows環境でのみ完全に機能します。**
  - 対処法: macOSやLinuxでは、ファイルは開きますが、特定のページやシートに自動的にジャンプしません。手動で移動する必要があります。

- **Word、Excel、PowerPointのページジャンプ機能を使用するには、アプリケーションのパスを設定する必要があります。**
  - 対処法: 上記の「Office アプリケーションのパス設定」セクションを参照して、各アプリケーションのパスを設定してください。

### Word文書ナビゲーションのトラブルシューティング

#### 「ActiveXオブジェクトを作成できません」というエラーが表示される

**解決策:**
1. コマンドプロンプトを管理者権限で実行してみてください
   - スタートメニュー → 「cmd」と入力 → 右クリック → 「管理者として実行」
2. セキュリティ設定でActiveXの使用が許可されているか確認してください

#### 「ファイルが見つかりません」というエラーが表示される

**解決策:**
1. ファイルパスが正しいか確認してください
2. ファイル名やパスに日本語や特殊文字が含まれている場合は、必ず引用符（"）で囲んでください
3. ファイルが実際に存在するか確認してください

#### 「コンパイルエラー」が表示される

**解決策:**
1. simple_navigator.js を使用してみてください
2. 最新版のスクリプトをダウンロードしてみてください

### 一般的なトラブルシューティング

#### ファイルが開かない場合

1. ファイルパスが正しいことを確認してください
2. ファイルが存在することを確認してください
3. ファイルが他のアプリケーションで開かれていないことを確認してください
4. VSCodeを再起動してみてください

#### ページジャンプが機能しない場合

1. リンク形式が正しいことを確認してください
2. Office アプリケーションのパスが正しく設定されていることを確認してください
3. Windows環境であることを確認してください（Word、Excel、PowerPointの場合）
4. VSCodeを管理者権限で実行してみてください
## リリースノート

### 1.0.0 (2023-10-01)

- 初回リリース
- PDFビューア機能の実装
- Word文書ナビゲーション機能の実装
- その他のOfficeファイル連携機能の実装

### 1.1.0 (2024-01-15)

- PDFビューアのページジャンプ機能を改善
- `#page=N`形式のページ指定に対応
- パフォーマンスの最適化

### 1.2.0 (2025-04-25)

- PDF以外のファイルタイプにもページジャンプ機能を追加
  - テキストファイルの行番号指定機能
  - Wordファイルのページ番号・ブックマーク指定機能
  - Excelファイルのシート名・セル位置指定機能
  - PowerPointファイルのスライド番号指定機能
- 各種リンク形式のサポートを拡充（`#page=N`, `?page=N`, `#N`など）
- Word文書ナビゲーション機能の強化
  - 段落構造の解析機能を追加
  - 特定の段落に即座にジャンプする機能を追加
  - コマンドラインツールの追加（simple_navigator.js, wordNavigator.js）
- READMEの大幅な改善と使用例の追加
- TypeScriptのコンパイラエラーを修正

## 今後の予定

ContractMate6は継続的に改善されています。今後のバージョンでは以下の機能を追加する予定です：

- **Word文書の構造解析**: 見出しや条項などの文書構造を解析し、ナビゲーションを容易にする機能
- **Excelデータの可視化**: VSCode内でExcelデータを表やグラフとして表示する機能
- **PowerPointのプレゼンテーションモード**: VSCode内でスライドショーを実行する機能
- **複数ファイルの一括検索**: 複数のPDFやOfficeファイル内のテキストを一括検索する機能
- **AIによる文書分析**: 契約書や法律文書の内容を分析し、重要なポイントを抽出する機能

---

## 開発者向け情報

### 拡張機能の構造

ContractMate6は、以下のコンポーネントで構成されています：

- **PDFハンドラー**: PDFファイルの表示とページジャンプ機能を提供
  - `pdfHandler.ts`: PDFファイルの処理とページジャンプ機能を実装
  - `PdfEditorProvider`: カスタムエディタプロバイダとしてPDFビューアを実装

- **テキストハンドラー**: テキストファイルの表示と行番号ジャンプ機能を提供
  - `textHandler.ts`: テキストファイルの処理と行番号ジャンプ機能を実装

- **Wordハンドラー**: Word文書の表示とページ・ブックマークジャンプ機能を提供
  - `wordHandler.ts`: Word文書の処理とページ・ブックマークジャンプ機能を実装

- **Excelハンドラー**: Excelファイルの表示とシート・セルジャンプ機能を提供
  - `excelHandler.ts`: Excelファイルの処理とシート・セルジャンプ機能を実装

- **PowerPointハンドラー**: PowerPointファイルの表示とスライドジャンプ機能を提供
  - `powerPointHandler.ts`: PowerPointファイルの処理とスライドジャンプ機能を実装

- **イメージハンドラー**: 画像ファイルの表示機能を提供
  - `imageHandler.ts`: 画像ファイルの処理と表示機能を実装

- **ユーティリティ**: 共通機能を提供
  - `fileUtils.ts`: ファイル操作や種類判定などの共通機能を実装

### ページジャンプ機能の実装

ページジャンプ機能は、以下の流れで実装されています：

1. URIからページ番号や位置情報を抽出
2. ファイルタイプに応じた適切なハンドラーを呼び出し
3. ハンドラーがファイルを開き、指定された位置にジャンプ

各ファイルタイプごとに異なるジャンプ方法を実装しています：

- **PDFファイル**: PDF.jsのAPIを使用してページを指定
- **テキストファイル**: VSCodeのエディタAPIを使用して行番号を指定
- **Wordファイル**: コマンドラインオプションを使用してページやブックマークを指定
- **Excelファイル**: コマンドラインオプションを使用してシートやセルを指定
- **PowerPointファイル**: コマンドラインオプションを使用してスライドを指定

### 拡張機能の開発

ContractMate6の開発に貢献するには、以下の手順に従ってください：

1. リポジトリをクローン
2. 依存関係をインストール: `npm install`
3. 拡張機能をビルド: `npm run compile`
4. デバッグ実行: F5キーを押すか、「実行とデバッグ」メニューから「拡張機能を起動」を選択

#### コーディング規約

- TypeScriptの型定義を適切に使用
- コメントは日本語で記述
- 関数やクラスには適切なJSDocコメントを付ける
- エラー処理を適切に実装

### テスト

拡張機能のテストには、`test.md`ファイルを使用します。このファイルには、各機能をテストするためのリンクとテスト手順が記載されています。

#### テスト手順

1. `test.md`ファイルを開く
2. 各テストケースのリンクをクリックして、機能が正しく動作することを確認
3. 問題がある場合は、コンソールログを確認して原因を特定

### トラブルシューティング

開発中に問題が発生した場合は、以下の手順を試してください：

1. VSCodeの出力パネルでログを確認
2. デバッガを使用してコードの実行を追跡
3. TypeScriptのコンパイラエラーを確認

**Enjoy!**
### 文書構造分析機能の詳細

Word文書ナビゲーション機能の中核となる「文書構造分析」機能について、詳細な説明と使い方を紹介します。

#### 機能概要

文書構造分析機能は、Word文書内のすべての段落、見出し、条項を自動的に識別し、その構造を解析します。この機能により、以下のことが可能になります：

- 文書内のすべての段落を一覧表示
- 見出しや条項を自動的に識別
- 段落のスタイル情報を取得
- 文書の階層構造（アウトラインレベル）を把握
- 特定の条項や見出しに即座にジャンプ

#### 使用方法

##### 基本的な使い方

1. コマンドプロンプトまたはPowerShellを開きます
2. 以下のコマンドを実行します：

```
cscript wordNavigator.js "C:\path\to\your\document.docx" 1
```

このコマンドを実行すると、以下の処理が行われます：
- 指定したWord文書が開かれます
- 文書内のすべての段落の構造が分析されます
- 分析結果がコンソールに表示されます
- 1段落目（通常はタイトルなど）に移動します

##### 出力例

```
段落番号: 1, スタイル: 見出し 1, アウトラインレベル: 1, 内容: サンプル契約書
段落番号: 2, スタイル: 標準, アウトラインレベル: 9, 内容: 
段落番号: 3, スタイル: 標準, アウトラインレベル: 9, 内容: 株式会社〇〇（以下「甲」という）と株式会社△△（以下「乙」という）とは、以下のとおり契約（以下「本契約」という）を締結する。
段落番号: 4, スタイル: 標準, アウトラインレベル: 9, 内容: 
段落番号: 5, スタイル: 見出し 2, アウトラインレベル: 2, 内容: 第1条（目的）
段落番号: 6, スタイル: 標準, アウトラインレベル: 9, 内容: 
段落番号: 7, スタイル: 標準, アウトラインレベル: 9, 内容: 本契約は、甲が乙に対して委託する業務の内容および条件を定めることを目的とする。
...
```

#### 出力情報の解釈

- **段落番号**: 文書内での段落の位置（1から始まる連番）
- **スタイル**: 段落に適用されているWordのスタイル名（「見出し 1」「標準」など）
- **アウトラインレベル**: 文書の階層構造を示す値（数値が小さいほど上位階層）
  - 1: 最上位の見出し（通常は「見出し 1」スタイル）
  - 2: 2番目の階層（通常は「見出し 2」スタイル）
  - ...
  - 9: 本文（通常は「標準」スタイル）
- **内容**: 段落のテキスト内容

#### 活用方法

##### 1. 特定の条項にジャンプする

出力された段落情報から、目的の条項の段落番号を特定し、そのまま移動できます：

```
# 出力から「第3条（契約期間）」が13段落目だと分かった場合
cscript wordNavigator.js "C:\path\to\your\document.docx" 13
```

##### 2. 文書の構造を把握する

出力された情報を確認することで、文書の全体構造を素早く把握できます：
- 見出しの階層関係
- 各条項の位置
- 文書のセクション構成

##### 3. 条項番号と段落番号のマッピングを作成

頻繁に参照する文書の場合、以下のようなマッピング表を作成しておくと便利です：

| 条項名 | 段落番号 |
|-------|---------|
| タイトル | 1 |
| 前文 | 3 |
| 第1条（目的） | 5 |
| 第2条（業務内容） | 9 |
| 第3条（契約期間） | 13 |
| ... | ... |

このマッピング表を使用することで、必要な条項に即座にジャンプできます。

#### 高度な使用方法

##### 特定のスタイルの段落のみを抽出

文書構造分析の結果をフィルタリングして、特定のスタイル（例えば見出し）のみを表示することもできます。これは、`testWordNavigator.js`を修正することで実現できます：

```javascript
// 見出しスタイルの段落のみを表示する例
for (var i = 0; i < structure.length; i++) {
    var para = structure[i];
    if (para.style.indexOf('見出し') !== -1) {
        WScript.Echo("見出し: " + para.index + ", " + para.text);
    }
}
```

##### 文書の目次を自動生成

文書構造分析の結果を使用して、文書の目次を自動生成することもできます：

```javascript
// 目次を生成する例
WScript.Echo("=== 目次 ===");
for (var i = 0; i < structure.length; i++) {
    var para = structure[i];
    if (para.outlineLevel < 9) {  // アウトラインレベルが9未満（見出し）の場合
        var indent = "";
        for (var j = 1; j < para.outlineLevel; j++) {
            indent += "  ";  // アウトラインレベルに応じてインデント
        }
        WScript.Echo(indent + para.text + " (段落番号: " + para.index + ")");
    }
}
```

この機能を活用することで、長い契約書や法律文書の操作効率が大幅に向上します。特に、クライアントとの打ち合わせ中に特定の条項をすぐに参照したい場合や、複数の文書間で同様の条項を比較したい場合に非常に便利です。