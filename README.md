# GAS_Quick-Parts
Quick Parts feature in GAS/Google Document.

# 使い方
## 導入
1. GoogleDocsにて2つドキュメントを作成。
    - 1つはパーツを格納
    - もう1つはスクリプトを格納
2. スクリプトを格納する方のドキュメントにて、「拡張機能 > AppsScript」、もしくは「ツール > スクリプトエディタ」 からスクリプトエディタを表示し、`script.gs` の内容を転記
3. この時、2行目の`const TemplateDoc = DocumentApp.openById('')` にパーツ格納用のドキュメントのID（https://docs.google.com/document/d/ **XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX** /edit の**X**の部分）を入れておく
4. エディタの左サイドバーから「トリガー」を選択し、右下「トリガーを追加」、ダイアログが表示されたら「保存」をクリック  
※ 一度スクリプトエディタから「実行」を実施して、権限を渡す必要がある可能性があります。
5. スクリプトを格納する方のドキュメントを更新し、「ヘルプ」の道に「Quick Parts」が表示されれば完了

## パーツの追加
1. パーツ格納用のドキュメントにて、`{parts_name}` と `/{parts_name}` で挟むように追加出来るようにしたいパーツを記載  
![](/img/part.png)
2. スクリプト用のドキュメントにて、スクリプトエディタを開き、21行目の `Template()` 関数をコピー&ペースト  
3. 任意の関数名と、 `QuickParts()` の引数に手順1で作成した `parts_name` を指定  
![](/img/foothold.png)
4. `onOpen()` 関数内で、任意の箇所に `.addItem('表示名', 'parts_name')` を追加  
![](/img/add_menu.png)
5. 保存した後、スクリプト用のドキュメントを更新し、手順4で指定した表示名の項目が追加されていれば完了  

## パーツの挿入
1. 「Quick Parts」メニューから挿入したいパーツ名を選択
