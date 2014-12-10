(defproject shinmuro/clj-poi "0.2.15"
  :description "thin wrapper of read/write xls format by apache POI."
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}

  :jvm-opts ["-XX:+TieredCompilation" "-XX:TieredStopAtLevel=1" "-Xverify:none"]

  :dependencies [[org.clojure/clojure "1.6.0"]
                 [shinmuro/util "0.4.0"]
                 [org.apache.poi/poi "3.11-beta3"]
                 [org.apache.poi/poi-ooxml "3.11-beta3"]
                 [org.apache.poi/poi-ooxml-schemas "3.11-beta3"]]

  :profiles {:dev {:dependencies [[com.taoensso/timbre "3.3.1"]]}})
