(ns clj-poi.xls
  "Apache POI の Clojure 向け簡易ラッパー。

   対応ファイルは xls (Excel 97-2003 フォーマット)のみ。
   Excel エラー値(#N/A など)も扱わない。
   エラー値が含まれる Excel を読み込んだ場合は例外が発生して終了する。"

  (:import  [org.apache.poi.ss.usermodel
             WorkbookFactory Workbook Sheet Row Cell CellStyle DateUtil]
            org.apache.poi.hssf.usermodel.HSSFWorkbook
            org.apache.poi.hssf.util.HSSFColor
            org.apache.poi.ss.util.CellRangeAddress

            java.util.TimeZone)

;  (:use     clj-util.core
  (:use     shinmuro.util
            [clojure.set :only [map-invert union difference]])
  (:require [clojure.java.io :as io]
            [clojure.string :as s]))
;; profile 実行時のみ有効にする
;            [taoensso.timbre :as timbre]))
;(timbre/refer-timbre)

(defn new-xls
  "空の xls 形式 Workbook を生成する。"
  []
  (HSSFWorkbook.))

(defn load-xls
  "指定パスの Excel を開く。"
  [file]
  (with-open [f (io/input-stream file)]
    (WorkbookFactory/create f)))

(defn save-xls!
  "Workbook オブジェクト wb を指定パス path に保存する。"
  [path ^Workbook wb]
  (with-open [out (io/output-stream path)]
    (.write wb out)))

(defn add-sheet!
  "Workbook オブジェクト wb に　sheet-name のシート名でシートを追加する。
   sheet-name の指定がない場合は \"Sheet0\" で名前を付与する。"
  ([^Workbook wb] (.createSheet wb))
  ([^Workbook wb sheet-name] (.createSheet wb sheet-name)))

(defn sheets
  "Workbook オブジェクト wb に存在する Sheet オブジェクトをシーケンスで返す。"
  [^Workbook wb]
  (map #(.getSheetAt wb %) (range (.getNumberOfSheets wb))))

(defn rows
  "Sheet オブジェクト st に存在する Row オブジェクトをシーケンスで返す。"
  [^Sheet st]
  (iterator-seq (.rowIterator st)))

(defn cols
  "Row オブジェクト row に存在する Cell オブジェクトをシーケンスで返す。"
  [^Row row]
  (iterator-seq (.cellIterator row)))

(defmulti cell!
  "cell に値 v を入れる。値の型に応じて書式も設定する。
   (主に POI API に日付書式自動判定入力が無い為)"
  (fn [_ v] (type v)))

(defmethod cell! nil [c v]
  (doto c
    (.setCellValue (str v))
    (.setCellType  Cell/CELL_TYPE_BLANK)))
(defmethod cell! java.lang.Number  [c v] (.setCellValue c (double v)) c)
(defmethod cell! java.lang.String  [c v] (.setCellValue c v) c)
(defmethod cell! java.lang.Boolean [c v] (.setCellValue c v) c)
(defmethod cell! java.util.Date    [c v] (.setCellValue c v) c)
(defmethod cell! :default          [c v]
  (try
    (do
      (.setCellValue c v)
      c)
    (catch Exception e (throw e))))

(defn- row
  "指定 Sheet オブジェクト st の index に該当する Row オブジェクトを返す。
   既に作成されていれば作成済み Row を、作成されていなければ指定 index の新規 Row オブジェクトを返す。"
  [^Sheet st index]
  (if-let [r (.getRow st index)]
    r
    (.createRow st index)))

(defn- cell
  "指定 Row オブジェクト row の index に該当する Cell オブジェクトを返す。
   既に作成されていれば作成済み Cell を、作成されていなければ指定 cellnum の新規 Cell オブジェクトを返す。"
  [^Row row index]
  (if-let [c (.getCell row index)]
    c
    (.createCell row index)))

(declare set-style!)
(defn row!
  "指定シート st の rowidx に coll の内容を追加する。

   マップデータによるスタイル指定可だが、
     - coll と style-maps の要素数は合わせる事
     - スタイル指定しないカラムは nil にする事。"
  ([st rowidx coll]
     (let [r (row st rowidx)]
       (mapall (fn [i value]
                 (when value (cell! (cell r i) value)))
               (range) coll)
       r))
  ([st rowidx coll style-maps]
     (if (not= (count coll) (count style-maps))
       (throw (IllegalStateException.
               (str "Item count of coll and style-maps is not equal, "
                    "coll is " (count coll) ", style-maps is " (count style-maps)))))
     (let [r (row st rowidx)]
       (mapall (fn [i value style]
                 (let [c (cell r i)]
                   (when value (cell! c value))
                   (when style (set-style! c style))))
               (range) coll style-maps)
       r)))

(defn write-seq!
  "Sheet オブジェクト st の先頭から 2 次元ベクタ nested-seq を書き込む。
   nested-seq のカラム数が異なる時は A 列から右へ埋められる。

   マップデータによるカラム単位のスタイル指定可だが、
     - nested-seq のカラム数と style-maps の要素数は合わせる事
     - スタイル指定しないカラムは nil にする事"
  ([st nested-seq]
     (doseq [[i e] (map-indexed vector nested-seq)]
       (row! st i e (repeat (count e) nil))))
  ([st nested-seq col-styles]
     (doseq [[i e] (map-indexed vector nested-seq)]
       (row! st i e col-styles))))

(defmulti cval
  "Excel cell 値読み取り。c に格納されてる値に応じた Clojure データを返す。"
  (fn [c] (.getCellType c)))

(defmethod cval Cell/CELL_TYPE_BLANK   [c] nil)
(defmethod cval Cell/CELL_TYPE_STRING  [c] (.getStringCellValue c))
(defmethod cval Cell/CELL_TYPE_FORMULA [c] (.getCellFormula c))
(defmethod cval Cell/CELL_TYPE_BOOLEAN [c] (.getBooleanCellValue c))

(defn as-date
  "数値セル c の値を日付として取り出す。"
  [^Cell c]
  (let [tz-offset (-> (System/getProperty "user.timezone")
                      TimeZone/getTimeZone
                      .getRawOffset)]
    (java.util.Date. (+ (.. c getDateCellValue getTime) tz-offset))))

(defmethod cval Cell/CELL_TYPE_NUMERIC
  [^Cell c]
  (if (DateUtil/isCellDateFormatted c)
    (as-date c)
    (.getNumericCellValue c)))

;; 未対応
(defmethod cval Cell/CELL_TYPE_ERROR [c]
  (throw (UnsupportedOperationException. "Excel cell エラー値あり。")))

(defn read-as-rs
  "Workbook オブジェクト wb のシート名 sheet-name 内にある値を
   clojure.java.jdbc ライクな形式で読み込む。

   [clojure.java.jdbc ライクな形式]
   sheet の一行目をフィールド名として扱い、二行目以降をデータとして
   フィールド名:値のマップデータのシーケンスとして読み込む。

   ついでにこれまた clojure.java.jdbc ライクにマップデータ後の関数として
   row-fn を渡せる。row 受け取るのはマップ化された row データ 1 行。
   デフォルトは identity, つまり as-is で返す。

   split-pos までを header 行、 split-pos より後を data 行で分ける。

   header-fn は上記で分割された header 行中の更に何行目をカラム列名とするかを指定する。
   first, second くらいしか想定してないが、関数渡しなので nth もやればいける？"
  [^Workbook wb ^String sheet-name & {:keys [row-fn split-pos header-fn] 
                                      :or   {row-fn identity
                                             split-pos 2
                                             header-fn second}}]
  (if-let [st (first (filter #(= sheet-name (.getSheetName %)) (sheets wb)))]
    (let [st-rows (rows st)
          max-col-index (apply max (map #(.getLastCellNum %) st-rows))
          sheet-data (for [row st-rows]
                       (for [i (range max-col-index)]
                         (cval (.getCell row i Row/CREATE_NULL_AS_BLANK))))
          [header data] (split-at split-pos sheet-data)
          header (map keyword (header-fn sheet-data))
          map-data (map (partial zipmap header) data)]
      (map row-fn map-data))
    (throw (IllegalArgumentException. (str "シート名 " sheet-name " が見つかりません。")))))

(defn- styles
  "Excel に設定された CellStyle オブジェクトセットを取得。"
  [^Workbook xls]
  (map #(.getCellStyleAt xls (short %))
       (range (.getNumCellStyles xls))))

(let [border-map {:none 0 :thin 1 :medium 2 :dashed 3 :dotted 4 :thick 5 :double 6 :hair 7
                  :medium-dashed 8 :dash_dot 9 :medium-dash-dot 10 :dash-dot-dot 11
                  :medium-dash-dot-dot 12 :slanted-dash-dot 13}]
  (defn- border-val [key]
    (let [cameled (dash->camel (name key))]
      {:getter (str "get" cameled) :setter (str "set" cameled)
       :val-fn border-map})))

(def ^:private color-index (into {} (HSSFColor/getIndexHash)))

(defn- colornum->key
  [n]
  (-> (color-index n) class .getSimpleName const->key))

(def color-map
  "Keyword 色名と Excel カラーインデックス値のマップデータ。"
  (-> (zipmap (map colornum->key (keys color-index))
              (keys color-index))
      (assoc :automatic 64
             :none 0)))

;; TODO: 場合によっては必要になるかもしれないもの
;;   :font-index (or :font) - index でしか取れない
;;   :index
;;   :parent-style                   ; 殆ど null らしい。xls のみ。多分不要。
(def ^:private convert-map
  "CellStyle オブジェクト変換マップ(暫定)

  {:prop-name {:getter String :setter String
               :val-fn map-or-fn}}"

  {:index          {:getter "getIndex" :val-fn identity} ; 取得のみ。POI でも setter API 無し。
   :halign         {:getter "getAlignment" :setter "setAlignment"
                    :val-fn {:general 0 :left 1 :center 2 :right 3
                             :fill 4 :justify 5 :center-selection 6}}
   :valign         {:getter "getVerticalAlignment" :setter "setVerticalAlignment"
                    :val-fn {:top 0 :center 1 :bottom 2 :justify 3}}
   :wrapped?       {:getter "getWrapText"  :setter "setWrapText" :val-fn identity}
   :shrink-to-fit? {:getter "getShrinkToFit" :setter "setShrinkToFit" :val-fn identity}
   :locked?        {:getter "getLocked" :setter "setLocked" :val-fn identity}
   :hidden?        {:getter "getHidden" :setter "setHidden" :val-fn identity}
   :format         {:getter "getDataFormatString"
                    :setter "setDataFormat"
                    :val-fn identity}
   :border-bottom  (border-val :border-bottom)
   :border-left    (border-val :border-left)
   :border-right   (border-val :border-right)
   :border-top     (border-val :border-top)

   :bottom-border-color {:getter "getBottomBorderColor" :setter "setBottomBorderColor"
                         :val-fn color-map}
   :left-border-color   {:getter "getLeftBorderColor" :setter "setLeftBorderColor"
                         :val-fn color-map}
   :right-border-color  {:getter "getRightBorderColor" :setter "setRightBorderColor"
                         :val-fn color-map}
   :top-border-color    {:getter "getTopBorderColor" :setter "setTopBorderColor"
                         :val-fn color-map}
   
   :fill-background-color {:getter "getFillBackgroundColor"
                           :setter "setFillBackgroundColor"
                           :val-fn color-map}
   :fill-foreground-color {:getter "getFillForegroundColor"
                           :setter "setFillForegroundColor"
                           :val-fn color-map}
   :fill-pattern {:getter "getFillPattern" :setter "setFillPattern"
                  :val-fn {:no-fill 0 :solid-foreground 1 :fine-dots 2 :alt-bars 3
                           :sparse-dots 4 :thick-horz-bands 5 :thick-vert-bands 6
                           :thick-backward-diag 7 :thick-forward-diag 8 :big-spots 9
                           :bricks 10 :thin-horz-bands 11 :thin-vert-bands 12
                           :thin-backward-diag 13 :thin-forward-diag 14 :squares 15
                           :diamonds 16 :less-dots 17 :least-dots 18}}
   :indention {:getter "getIndention" :setter "setIndention" :val-fn identity}
   :rotation  {:getter "getRotation" :setter "setRotation" :val-fn identity}})

(defn- parent-workbook
  [^Cell c]
  (.. c getSheet getWorkbook))

(defn- create-format
  [^Cell c]
  (let [wb (parent-workbook c)]
    (.. wb getCreationHelper createDataFormat)))

(defn style->map
  "CellStyle オブジェクトをマップデータ化したものを返す。Font 未対応。"
  [^CellStyle style]
  (reduce (fn [m [k v]]
            (let [jvm-val (clj-invoke style (:getter v))
                  val-fn (if (fn? (:val-fn v)) (:val-fn v) (map-invert (:val-fn v)))]
              (assoc m k (val-fn jvm-val))))
          {} convert-map))

(defn- style-set
  [^Workbook xls]
  (set
   (map (comp #(dissoc %1 :index)
              style->map)
        (styles xls))))

(let [latest-wb        (atom nil)
      latest-style-num (atom 0)
      latest-styles    (atom nil)]
  (defn- all-styles
    [^Workbook xls]
    (when (not= xls @latest-wb)
      (reset! latest-wb nil)
      (reset! latest-style-num 0)
      (reset! latest-styles nil))
    (if (> (.getNumCellStyles xls) @latest-style-num)
      (let [jvm-styles (styles xls)
            clj-styles (map (comp #(dissoc %1 :index)
                                  style->map)
                            jvm-styles)]
        (reset! latest-wb xls)
        (reset! latest-style-num (.getNumCellStyles xls))
        (reset! latest-styles (zipmap clj-styles jvm-styles)))
      @latest-styles)))

(comment
  ;; REPL でテスト用に Workbook, Cell とかを作るコード
  (def xls (new-xls))
  (def st (add-sheet! xls))
  (def r1 (row! st 0 ["a" "b" "c"]))
  (def c1 (-> st rows first cols first))

  (def xls (new-xls))
  (def st (add-sheet! xls))
  (def coll [["a" "b" 10 20 30 nil "d" (java.util.Date.)]
             ["d" "e" 40 50 60 nil "f" (java.util.Date.)]])
  (def colstyles [nil nil nil nil nil nil nil {:format "yyyy/mm/dd hh:mm:ss"}])
  (write-seq! st coll colstyles)

  (save-xls! "a.xls" xls)

  ;; style 設定パフォーマンス検証
  (def yellow-bg {:fill-pattern :solid-foreground
                  :fill-foreground-color :yellow})

  (time (set-style! c1 yellow-bg))
  "Elapsed time: 17.924668 msecs"

  (profile :info :set-style! (dotimes [_ 100] (set-style! c1 yellow-bg)))
  ;; set-style! は単体で見るとさほどではないが積み重なった場合やっぱり重い処理。
)

(defn- merge-new-style-fn
  [^CellStyle style style-map]
  (-> (dissoc (style->map style) :index)
      (merge style-map)))
(def ^:private merge-new-style (memoize merge-new-style-fn))

;; profile 実行版。
#_(defn set-style!
  [^Cell c style-map]
  (when style-map
    (let [wb    (p :wb (parent-workbook c))
          style (p :style (.getCellStyle c))
          maybe-new (p :maybe-new (-> (dissoc (style->map style) :index)
                                      (merge style-map)))
;          maybe-new (p :maybe-new (merge-new-style style style-map))
          all (p :all-styles (all-styles wb))
          all-smap (p :all-smap (set (keys all)))
          added-set (p :added-set (union all-smap #{maybe-new}))]
      (p :body
         (if (seq (difference added-set all-smap))
           (do
             (let [^CellStyle new-style (.createCellStyle wb)]
               (doseq [[k v] maybe-new]
                 (when-let [setter (:setter (convert-map k))]
                   (if (= k :format)
                     (clj-invoke new-style setter (.getFormat (create-format c) v))
                     (clj-invoke new-style setter
                                 (if (keyword? v)
                                   (-> convert-map k :val-fn v)
                                   v)))))
               (.setCellStyle c new-style)))
           (.setCellStyle c (all maybe-new)))))))

(defn set-style!
  "指定セルに各スタイルをセットする(Mutable)。

   Excel はスタイル数制限が結構厳しい為、この関数ではマップデータ化した既存 CellStyle
   全てと比較して、存在してない場合に初めて生成するようにしている。

   スタイル数の上限チェックはしていない。"
  [^Cell c style-map]
  (when style-map
    (let [wb    (parent-workbook c)
          style (.getCellStyle c)
          maybe-new (merge-new-style style style-map)
          all (all-styles wb)
          all-smap (set (keys all))
          added-set (union all-smap #{maybe-new})]
      (if (seq (difference added-set all-smap))
        (do
             (let [^CellStyle new-style (.createCellStyle wb)]
               (doseq [[k v] maybe-new]
                 (when-let [setter (:setter (convert-map k))]
                   (if (= k :format)
                     (clj-invoke new-style setter (.getFormat (create-format c) v))
                     (clj-invoke new-style setter
                                 (if (keyword? v)
                                   (-> convert-map k :val-fn v)
                                   v)))))
               (.setCellStyle c new-style)))
        (.setCellStyle c (all maybe-new))))))

(defn color-sheet
  "Clojure 色名キーワードと実際の色を出力。
     ファイル名: poi-clj-color-sheet.xls。"
  []
  (let [xls (new-xls)
        st  (add-sheet! xls "color sheet")]
    (maprun
      (fn [i [k v]]
        (row! st i [(name k) nil] [nil {:fill-pattern :solid-foreground
                                        :fill-foreground-color v}]))
      (range (count color-map))
      color-map)
    (save-xls! "poi-clj-color-sheet.xls" xls)))

(defn set-auto-filter
  "指定範囲 range-str でのオートフィルタを有効にする。
   range-str には Excel で馴染みの形式を使用可。"
  [st range-str]
  (.setAutoFilter st (CellRangeAddress/valueOf range-str)))
