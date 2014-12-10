(ns poi-clj.excel-test
  (:require [clojure.test :refer :all]
            [poi-clj.excel :refer :all]))

#_(deftest current-dir-check
  (println "current dir: " (System/getProperty "user.dir")))

;; 日付NG。
;; 読取り時に表示書式が日付になってないと判別のしようがないが
;; その前に CellStyle をなるべく増やさないような設定の仕方を考える必要がある。
(deftest test-cell!
  (testing "cell への型に応じた値の入力確認"
    (let [wb (new-xls)
          st (add-sheet! wb)
          row (row! st 0 [1])
          cell (first row)]
      (are [expected expr] (= expected expr)
           nil  (cval (cell! cell nil))
           10.0 (cval (cell! cell 10))
           "a"  (cval (cell! cell "a"))
           #inst"2014-08-01T05:00:00.000-00:00"
           (cval (cell! cell #inst"2014-08-01T05:00:00.000-00:00"))))))

(deftest test-create-row!-if-contains-nil
  (testing
      "row! で呼んでる cell! で nil な値を渡すと NPE になっていたので対処できてるかどうかのテスト"
    (let [wb (new-xls)
          st (add-sheet! wb)
          data ["a" nil "c"]]
      (is (= data (map cval (cols (row! st 0 data))))))))

;; [ok] 読込一連のテスト
(comment
#_(def xls (load-excel "format_check.xls"))
#_(def st (first (sheets xls)))
;; 空白 Cell はまるきり飛ばされてた
#_(def all (mapcat cols (rows st)))

#_(def dtc (nth all 1))

;; write => OK
#_(with-open [wtr (io/output-stream "row_65535_copy.xls")]
  (.write xls wtr))
)

;; [ok] 書出一連のテスト
(comment
  (def out-wb (create-xls))

  (create-sheet! out-wb)

  (def st (first (sheets out-wb)))

  (def test-vec
    [["a" "b"]
     ["d" "e" "f"]])

  ;; ここで keep-indexed の戻り値は不要
  (write-seq! st test-vec)
  (save-excel! "out6.xls" out-wb)
)

;; [ok] 大量書出しテスト(65535)
(comment
  (def out-wb (create-xls))
  (def st (create-sheet! out-wb))

  (def atoz (map (comp str char)
                 (range (int \A) (int (int \Z)))))
  (def large-vec (repeat 65535 atoz))

  (write-seq! st large-vec)

  (save-excel! "large-out.xls" out-wb)
  )
