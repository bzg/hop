(ns bzg.hop-test
  (:require [babashka.process :refer [shell]]
            [clojure.string :as str]
            [babashka.fs :as fs]))

(def parser "src/bzg/hop.clj")
(def test-org "test/bzg/test.org")
(def expected-dir "test/bzg/expected")

(def formats
  [["md" "md"] ["html" "html"] ["org" "org"]
   ["json" "json"] ["edn" "edn"] ["yaml" "yaml"] ["ics" "ics"]])

(defn- expected-file [ext] (str expected-dir "/test." ext))

(defn- run-parser [fmt]
  (:out (shell {:out :string :err :string} "bb" parser "-f" fmt test-org)))

(defn- normalize [ext text]
  (if (= ext "ics")
    (str/replace text #"DTSTAMP:\d{8}T\d{6}" "DTSTAMP:NORMALIZED")
    text))

(defn generate []
  (fs/create-dirs expected-dir)
  (doseq [[fmt ext] formats]
    (let [output (run-parser fmt)
          path   (expected-file ext)]
      (spit path output)
      (println (str "  wrote " path " (" (count (str/split-lines output)) " lines)"))))
  (println "Done. Review test/bzg/expected/ before committing."))

(defn test-all []
  (let [results
        (for [[fmt ext] formats]
          (let [path (expected-file ext)]
            (if-not (fs/exists? path)
              (do (println (str "SKIP " fmt " — run: bb test:generate")) :skip)
              (let [actual   (run-parser fmt)
                    expected (slurp path)]
                (if (= (normalize ext actual) (normalize ext expected))
                  (do (println (str "  OK " fmt)) :pass)
                  (let [actual-path (str path ".actual")]
                    (spit actual-path actual)
                    (println (str "FAIL " fmt " → diff " path " " actual-path))
                    :fail))))))]
    (let [{:keys [pass fail skip]} (merge {:pass 0 :fail 0 :skip 0} (frequencies results))]
      (println (str "\n" pass "/" (count results) " passed"
                    (when (pos? fail) (str ", " fail " failed"))
                    (when (pos? skip) (str ", " skip " skipped"))))
      (when (pos? fail) (System/exit 1)))))
