# 弁護士業務支援用 Word文書ナビゲーションツール

## 概要
このツールは、弁護士業務の効率化を目的としたWord文書操作ツールです。Windows Script Host (WSH)とActiveXObjectを使用して、Word文書を開き、文書内の段落構造を解析し、特定の段落に即座に移動することができます。

## 主な機能
- ✅ Word文書を即座に開く
- ✅ 文書の全段落構造を取得（条項や見出しも判別可能）
- ✅ 特定の条項や段落番号を指定し、即座にその場所にジャンプ

## 必要環境
- Windows OS
- Microsoft Word
- Windows Script Host (WSH)

## 使用方法
コマンドプロンプトまたはPowerShellから以下のコマンドを実行します：

```
cscript wordNavigator.js [Word文書パス] [移動したい段落番号]
```

### 引数の説明
1. `[Word文書パス]` - 開きたいWord文書のパス（.docまたは.docx）
2. `[移動したい段落番号]` - ジャンプしたい段落の番号（1から始まる整数）

### 実行例
```
cscript wordNavigator.js "C:\契約書\賃貸契約書.docx" 12
```

この例では、「賃貸契約書.docx」を開き、12番目の段落に自動的にジャンプします。

## 出力情報
実行すると、文書内の全段落の情報が表示されます：
- 段落番号
- スタイル名
- アウトラインレベル
- 段落の内容（テキスト）

これにより、文書の構造を把握し、必要な段落を特定することができます。

## 活用例
- 契約書の特定条項を即座に参照する
- クライアントとの打合せ中に必要な条文を素早く提示する
- 長文の法律文書内を効率的に移動する
- 複数の文書間で同様の条項を比較検討する
- 文書の検索・確認作業の時間を大幅に短縮する

## 注意事項
- このスクリプトを実行するには、セキュリティ設定でActiveXObjectの使用が許可されている必要があります。
- 初回実行時にセキュリティ警告が表示される場合があります。