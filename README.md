clj-poi.xls
================================================================================

**JAPANESE DOCUMENTATION ONLY**

apache POI thin wrapper only xls format read/write.

自分用 apache POI ラッパー。

## 特徴
- xls 形式での書式の最大数が 4000 と言う事なので、なるべく使い回すようにしてます
- 書き出すのがそれなりに速いかもしれません
- 上記 Excel 仕様が特に気にならない用途であれば docjure などの方が使いやすいかもしれません
- 限定的ながら書式設定も一応できます
- clojure.java.jdbc が吐き出す resultset ライクなデータで読み込む関数が一応あります(read-as-rs)

## 制限事項
- 旧 Excel 形式(xls)にしか対応してません
- 日付型で対応してるのは java.util.Date ですが、ローカル時間に無理やり変換してます
- 書式設定でフォントは未対応です

## 使い方
~~``project.clj`` の ``:dependencies`` に以下を追加して ``lein deps``~~
```clojure
~~[poi-clj "0.x.x"]~~
```

Clojars に deploy してませんのでしばしお待ち下さい。

ローカルで構わなければ[shinmuro/util](https://github.com/shinmuro/util)とここのリポジトリ
をそれぞれ git clone してlein install すれば使用可能です。

後は使いたい所で
```clojure
(require '[shinmuro.clj-poi.xls :as xl])
```
など。

各関数の使い方は[API doc](https://shinmuro.github.io/clj-poi/doc)見て何となく察して下さい。

## License

Copyright © 2014 shinmuro

Distributed under the Eclipse Public License either version 1.0. Same as Clojure.
