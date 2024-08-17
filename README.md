# pptx2aituberkitscript

このプログラムは、PowerPointのpptxをtegnikeさんのAITuberKitのスライドモードに対応したデータに変換します。pptxの各スライドのノートに記述した内容をアバターがしゃべるようにscripts.jsonを出力します。

[ssine/pptx2md](https://github.com/ssine/pptx2md)を改造しています。

## 使用方法
```
python -m pptx2aituberkitscript <pptxスライドのパス>
```

pptxスライド名から、pptx拡張子を除いたフォルダが作成されるので、そのフォルダをAITuberKitのpublic/slidesに移動させて、AITuberKitを起動して、設定画面からスライドモードをONにすると、先ほど移動させたスライド名が表示されるので、それを選択すると表示されるようになります。

tegnikeさんが作成したtheme.cssを一緒に生成しています。

現状、レイアウトとしてはpptxの「スライド タイトル」のみ対応しています（日本語PowerPointで生成した場合のレイアウト名）。

## 制限
・テーブルは変換できますが、AITuberKitでは表として表示されません。
・画像のサイズ調整等は後からする必要があります。
・スライドマスターの画像などは変換されません。
