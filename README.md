# label_gnerator

大昔に書いたラベル生成コードを手直ししてみた．地学サークルの標本展示，ご自宅の化石標本整理にどうぞ．
![fig](https://github.com/ammokun/label_generator/blob/master/figure1.png)

# 使い方
1. sample.xlsxに情報を書き込む．斜体にしたりしなかったりするので属名，種小名，cf.などの記号は別々のセルに入れる．使わない情報も混ざってるのは仕様．

2. python lavel_1.pyで実行．label_row，label_columnで1ページに書き込む行数，列数を指定可能．デフォルトの文字の大きさだとlabel_row=2，label_column=6がちょうどいい．python-pptxとopenpyxlを使うので入れてない場合pip install python-pptxとかで導入．test.pptxにラベルが書き出されるが，pptxファイルを開いたままだとエラー吐くので閉じてから実行すること．

3. 印刷して，（お好みに応じてラミネートして）細い線に沿って切る
