(ns bzg.hop-test
  (:require [babashka.process :refer [shell]]
            [clojure.string :as str]
            [clojure.test :as test]
            [babashka.fs :as fs]))

(def parser "src/bzg/hop.clj")
(def test-org "test/bzg/test.org")
(def expected-dir "test/bzg/expected")

(def formats
  [["md" "md"] ["html" "html"] ["org" "org"]
   ["json" "json"] ["edn" "edn"] ["yaml" "yaml"]
   ["ics" "ics" ["-z" "Europe/Paris"]]
   ["ics" "windowed.ics" ["-R" "2025-03-01" "-w" "12"
                          "-z" "Europe/Paris" "-d" "45" "-a"]]
   ["ics-anon" "anon.ics" ["-R" "2025-03-01" "-w" "12"
                           "-z" "Europe/Paris" "-d" "60" "-a"]]
   ["ics-anon" "unbounded.anon.ics" ["-z" "Europe/Paris"]]])

(defn- expected-file [ext] (str expected-dir "/test." ext))

(defn- run-parser [fmt args]
  (:out (apply shell {:out :string :err :string}
               "bb" parser "-f" fmt (concat args [test-org]))))

(defn- normalize [ext text]
  (if (str/ends-with? ext "ics")
    (str/replace text #"DTSTAMP:\d{8}T\d{6}Z" "DTSTAMP:NORMALIZED")
    text))

(defn- run-unit-tests
  "Run the clojure.test suite in bzg.hop-unit-test; :pass or :fail."
  []
  (require 'bzg.hop-unit-test)
  (let [{:keys [fail error]} (test/run-tests 'bzg.hop-unit-test)]
    (if (zero? (+ fail error))
      (do (println "  OK unit tests") :pass)
      (do (println "FAIL unit tests") :fail))))

(defn- test-golden
  "Compare one format's output with its expected file; :pass, :fail or :skip."
  [[fmt ext args]]
  (let [path (expected-file ext)]
    (if-not (fs/exists? path)
      (do (println (str "SKIP " fmt " — run: bb test:generate")) :skip)
      (let [actual   (run-parser fmt args)
            expected (slurp path)]
        (if (= (normalize ext actual) (normalize ext expected))
          (do (println (str "  OK " fmt " (" ext ")")) :pass)
          (let [actual-path (str path ".actual")]
            (spit actual-path actual)
            (println (str "FAIL " fmt " → diff " path " " actual-path))
            :fail))))))

(defn generate []
  (fs/create-dirs expected-dir)
  (doseq [[fmt ext args] formats]
    (let [output (run-parser fmt args)
          path   (expected-file ext)]
      (spit path output)
      (println (str "  wrote " path " (" (count (str/split-lines output)) " lines)"))))
  (println "Done. Review test/bzg/expected/ before committing."))

(defn test-all []
  (let [results                  (conj (mapv test-golden formats) (run-unit-tests))
        {:keys [pass fail skip]} (merge {:pass 0 :fail 0 :skip 0} (frequencies results))]
    (println (str "\n" pass "/" (count results) " passed"
                  (when (pos? fail) (str ", " fail " failed"))
                  (when (pos? skip) (str ", " skip " skipped"))))
    (when (pos? fail) (System/exit 1))))
