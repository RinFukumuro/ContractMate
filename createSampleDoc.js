// サンプルWord文書を自動生成するスクリプト

// ActiveXObjectを使用してWordアプリケーションを起動
try {
    var word = new ActiveXObject("Word.Application");
    word.Visible = true;
    
    // 新規文書を作成
    var doc = word.Documents.Add();
    
    // タイトルを追加
    word.Selection.Style = word.ActiveDocument.Styles("見出し 1");
    word.Selection.TypeText("サンプル契約書");
    word.Selection.TypeParagraph();
    
    // 前文
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("株式会社〇〇（以下「甲」という）と株式会社△△（以下「乙」という）とは、以下のとおり契約（以下「本契約」という）を締結する。");
    word.Selection.TypeParagraph();
    
    // 第1条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第1条（目的）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("本契約は、甲が乙に対して委託する業務の内容および条件を定めることを目的とする。");
    word.Selection.TypeParagraph();
    
    // 第2条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第2条（業務内容）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("乙は、甲の指示に基づき、以下の業務を行うものとする。");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("(1) システム開発業務");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("(2) システム保守業務");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("(3) その他甲が指示する業務");
    word.Selection.TypeParagraph();
    
    // 第3条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第3条（契約期間）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("本契約の有効期間は、契約締結日から1年間とする。ただし、期間満了の1ヶ月前までに甲乙いずれからも書面による別段の意思表示がないときは、同一条件にてさらに1年間継続するものとし、以後も同様とする。");
    word.Selection.TypeParagraph();
    
    // 第4条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第4条（報酬）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("1. 甲は、乙に対し、本契約に基づく業務の対価として、別途甲乙間で合意した金額を支払うものとする。");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("2. 報酬の支払いは、乙の請求に基づき、請求書受領月の翌月末日までに、乙の指定する銀行口座に振り込む方法により行うものとする。");
    word.Selection.TypeParagraph();
    
    // 第5条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第5条（機密保持）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("1. 甲および乙は、本契約の履行に関連して知り得た相手方の技術上、営業上その他業務上の情報（以下「機密情報」という）を、相手方の書面による事前の承諾なくして第三者に開示または漏洩してはならない。");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("2. 前項の規定は、本契約終了後も3年間効力を有するものとする。");
    word.Selection.TypeParagraph();
    
    // 第6条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第6条（契約解除）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("甲または乙は、相手方が本契約に違反し、相当の期間を定めて催告したにもかかわらず、当該期間内に違反が是正されないときは、本契約を解除することができる。");
    word.Selection.TypeParagraph();
    
    // 第7条
    word.Selection.Style = word.ActiveDocument.Styles("見出し 2");
    word.Selection.TypeText("第7条（協議事項）");
    word.Selection.TypeParagraph();
    
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("本契約に定めのない事項または本契約の解釈に疑義が生じた場合は、甲乙誠意をもって協議の上解決するものとする。");
    word.Selection.TypeParagraph();
    
    // 署名欄
    word.Selection.Style = word.ActiveDocument.Styles("標準");
    word.Selection.TypeText("以上、本契約の成立を証するため、本書2通を作成し、甲乙記名押印の上、各1通を保有する。");
    word.Selection.TypeParagraph();
    word.Selection.TypeParagraph();
    
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var day = date.getDate();
    
    word.Selection.TypeText(year + "年" + month + "月" + day + "日");
    word.Selection.TypeParagraph();
    word.Selection.TypeParagraph();
    
    word.Selection.TypeText("甲：株式会社〇〇");
    word.Selection.TypeParagraph();
    word.Selection.TypeText("   代表取締役 〇〇 〇〇  印");
    word.Selection.TypeParagraph();
    word.Selection.TypeParagraph();
    
    word.Selection.TypeText("乙：株式会社△△");
    word.Selection.TypeParagraph();
    word.Selection.TypeText("   代表取締役 △△ △△  印");
    
    // ファイルを保存
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var currentDir = fso.GetAbsolutePathName(".");
    var filePath = currentDir + "\\サンプル契約書.docx";
    
    doc.SaveAs(filePath);
    
    WScript.Echo("サンプル契約書を作成しました: " + filePath);
    
    // Wordを閉じる（オプション）
    // doc.Close();
    // word.Quit();
    
} catch (error) {
    WScript.Echo("エラーが発生しました: " + error);
}