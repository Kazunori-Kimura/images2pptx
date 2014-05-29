generate PPTX from Image Files
================================

画像ファイル (jpg, png, gif, bmp) から PowerPointファイルを生成します。

* 指定されたフォルダにある画像を中央に貼り付けたスライドを自動的に作成します。
  - 画像はスライドの70%のサイズに拡大 or 縮小されます。
  - 縦横比はスライドと同じ(3:4)になります。
* 大量のスクリーンショットを元にスライドを作成するような用途を想定しています。

usage
------

    images2pptx <image folder> <powerpoint file>

    例:
        images2pptx ~/Desktop/ScreenShot slide.pptx


