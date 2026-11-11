(ns bzg.hop-unit-test
  "Unit tests for hop's pure helpers and for the org round-trip, which the
   golden-file tests in bzg.hop-test do not cover."
  (:require [clojure.test :refer [deftest is testing]]
            [clojure.string :as str]
            [clojure.walk :as walk]
            [clojure.tools.cli :as cli]
            [bzg.organ :as organ]
            [bzg.hop :as hop]))

;; Private helpers under test
(def ^:private ics-fold-line         #'hop/ics-fold-line)
(def ^:private heading-to-slug       #'hop/heading-to-slug)
(def ^:private parse-options-string  #'hop/parse-options-string)
(def ^:private normalize-toc-levels  #'hop/normalize-toc-levels)
(def ^:private render-toc            #'hop/render-toc)
(def ^:private iso-to-ics-datetime   #'hop/iso-to-ics-datetime)
(def ^:private iso->org-timestamp    #'hop/iso->org-timestamp)
(def ^:private item-in-window?       #'hop/item-in-window?)
(def ^:private merge-intervals-zoned #'hop/merge-intervals-zoned)
(def ^:private prune-done-sections   #'hop/prune-done-sections)
(def ^:private resolve-link-href     #'hop/resolve-link-href)
(def ^:private render-inline         #'hop/render-inline)
(def ^:private cli-options           @#'hop/cli-options)

(defn- ldt [s] (java.time.LocalDateTime/parse s))
(defn- zdt [s] (.atZone (ldt s) (java.time.ZoneId/of "UTC")))
(defn- occurrences [s re] (count (re-seq re s)))

(deftest ics-fold-line-test
  (testing "short lines are untouched"
    (is (= "SUMMARY:x" (ics-fold-line "SUMMARY:x"))))
  (testing "long lines fold at 75 octets, continuations at 74 + leading space"
    (let [folded (ics-fold-line (apply str (repeat 200 "a")))
          lines  (str/split folded #"\r\n")]
      (is (= 75 (count (first lines))))
      (is (every? #(str/starts-with? % " ") (rest lines)))
      (is (every? #(<= (count (.getBytes ^String % "UTF-8")) 75) lines))
      (is (= 200 (count (str/replace folded #"\r\n " ""))))))
  (testing "multi-byte characters are never split"
    (let [text   (apply str (repeat 60 "é"))
          folded (ics-fold-line text)]
      (is (= text (str/replace folded #"\r\n " "")))
      (is (every? #(<= (count (.getBytes ^String % "UTF-8")) 75)
                  (str/split folded #"\r\n"))))))

(deftest heading-to-slug-test
  (is (= "hello-world" (heading-to-slug "Hello, World!")))
  (is (= "section" (heading-to-slug "!!!")))
  (is (= "bold-title" (heading-to-slug [{:type :bold :children [{:type :text :value "Bold"}]}
                                        {:type :text :value " title"}]))))

(deftest parse-options-string-test
  (is (nil? (parse-options-string nil)))
  (is (nil? (parse-options-string "  ")))
  (is (= {:toc true :num false :h 3 :x "foo"}
         (parse-options-string "toc:t num:nil H:3 x:foo"))))

(deftest toc-test
  (testing "level jumps are clamped to +1"
    (is (= [1 2 2 1] (map :level (normalize-toc-levels [{:level 1} {:level 3} {:level 2} {:level 1}])))))
  (testing "HTML TOC stays balanced across level jumps"
    (let [entries [{:level 1 :title "A" :id "a"} {:level 3 :title "B" :id "b"}
                   {:level 2 :title "C" :id "c"} {:level 1 :title "D" :id "d"}]
          html    (render-toc entries :html)]
      (is (= (occurrences html #"<ul>") (occurrences html #"</ul>")))
      (is (= (occurrences html #"<li>") (occurrences html #"</li>")))
      (is (str/starts-with? html "<nav id=\"table-of-contents\">"))
      (is (str/ends-with? html "</nav>")))))

(deftest ics-datetime-test
  (is (= {:dtstart "20250115T103000"} (iso-to-ics-datetime "2025-01-15T10:30")))
  (is (= {:dtstart "20250115"} (iso-to-ics-datetime "2025-01-15")))
  (is (= {:dtstart "20250115T103000" :dtend "20250115T113000"}
         (iso-to-ics-datetime "2025-01-15T10:30/2025-01-15T11:30"))))

(deftest org-timestamp-test
  (is (= "<2025-01-15>" (iso->org-timestamp "2025-01-15" nil)))
  (is (= "<2025-01-15 10:30 +1w>" (iso->org-timestamp "2025-01-15T10:30" "+1w")))
  (is (= "<2025-01-15 10:30-11:30>" (iso->org-timestamp "2025-01-15T10:30/2025-01-15T11:30" nil)))
  (is (= "<2025-01-15>--<2025-01-17>" (iso->org-timestamp "2025-01-15/2025-01-17" nil))))

(deftest item-in-window-test
  (let [win-s (ldt "2025-03-01T00:00") win-e (ldt "2025-03-08T00:00")]
    (is (item-in-window? {:start "2025-03-03T10:00"} win-s win-e))
    (is (not (item-in-window? {:start "2025-02-03T10:00"} win-s win-e)))
    (is (not (item-in-window? {:start "2025-03-08T00:00"} win-s win-e)) "upper bound is exclusive")
    (is (item-in-window? {:start "2025-03-08T00:00"} win-s nil) "nil win-e means unbounded")
    (testing "a repeater started years ago reaches the window"
      (is (item-in-window? {:start "2010-01-04T09:00" :repeater {:n 1 :unit :week}} win-s win-e))
      (is (not (item-in-window? {:start "2010-01-04T09:00" :repeater {:n 1 :unit :year}} win-s win-e))))))

(deftest merge-intervals-test
  (let [a [(zdt "2025-03-01T10:00") (zdt "2025-03-01T11:00")]
        b [(zdt "2025-03-01T10:30") (zdt "2025-03-01T12:00")]
        c [(zdt "2025-03-01T12:00") (zdt "2025-03-01T13:00")]]
    (is (= [[(first a) (second b)] c] (merge-intervals-zoned [a b c]))
        "overlapping intervals merge, touching ones stay distinct")))

(deftest prune-done-sections-test
  (let [ast (organ/parse-org "* TODO Keep\n** DONE Drop\ntext\n* CANX Also drop\n* Plain"
                             {:todo-keywords {:done ["CANX"]}})
        titles (fn [n] (->> (tree-seq map? :children n)
                            (filter #(= :section (:type %)))
                            (map #(organ/inline-text (:title %)))))]
    (is (= ["Keep" "Drop" "Also drop" "Plain"] (titles ast)))
    (is (= ["Keep" "Plain"] (titles (prune-done-sections ast #{"DONE" "CANX"}))))))

(deftest resolve-link-href-test
  (let [node {:link-type :file :target "notes.org" :url "file:notes.org"}]
    (is (= "notes.html" (resolve-link-href node :html)))
    (is (= "file:notes.org" (resolve-link-href node :org)))
    (binding [hop/*base-url* "https://base.io/"]
      (is (= "https://base.io/notes.html" (resolve-link-href node :md)))
      (is (= "#x" (resolve-link-href {:link-type :custom-id :target "x"} :md)) "anchors are left alone"))))

(deftest inline-markup-test
  (let [nodes (organ/parse-inline "*b* /i/ _u_ +s+ ~c~ =v= $x^2$ \\(a\\)")]
    (is (= "*b* /i/ _u_ +s+ ~c~ =v= $x^2$ \\(a\\)" (render-inline nodes :org)))
    (is (= "**b** *i* _u_ ~~s~~ `c` `v` $x^2$ \\(a\\)" (render-inline nodes :md)))
    (is (= "<strong>b</strong> <em>i</em> <u>u</u> <del>s</del> <code>c</code> <code>v</code> $x^2$ \\(a\\)"
           (render-inline nodes :html)))))

(defn- structure
  "AST stripped of line numbers and blank-line bookkeeping, which an
   export normalises, for round-trip comparison."
  [ast]
  (walk/postwalk #(if (map? %) (dissoc % :line :blank-lines-before :blank-lines-after-title) %)
                 (organ/clean-node ast)))

(deftest org-roundtrip-test
  (let [src   (str "#+TITLE: Round trip\n#+OPTIONS: toc:t num:t\n"
                   "#+TODO: TODO NEXT | DONE CANX\n#+STARTUP: overview\n#+HTML_HEAD: <a>\n#+HTML_HEAD: <b>\n\n"
                   "* NEXT [#A] Task :tag:\nSCHEDULED: <2025-01-15 Wed 10:00-11:00 +1w>\n"
                   ":PROPERTIES:\n:CUSTOM_ID: task\n:END:\n:LOGBOOK:\n- Note\n:END:\n"
                   "#+PROPERTY: header-args :results silent\n"
                   "Text with $x$ and \\(y\\).\n\n"
                   "\\begin{equation}\nE = mc^2\n\\end{equation}\n\n"
                   "- item\n  | a | b |\n  |---+---|\n  | 1 | 2 |\n"
                   "- item with more\n  #+FOO: bar\n  :NOTES:\n  hello\n  :END:\n\n"
                   "  second paragraph\n  - nested\n"
                   "1. one\n\n   continued\n2. two\n\n"
                   "#+NAME: src\n#+BEGIN_SRC clojure -n\n(+ 1 2)\n#+END_SRC\n\n"
                   "#+BEGIN_QUOTE\nFirst.\n\nSecond.\n#+END_QUOTE\n"
                   "* DONE Closed\nCLOSED: [2025-01-16 Thu 09:00]\n")
        once  (hop/render-ast-as-org (organ/parse-org src))
        twice (hop/render-ast-as-org (organ/parse-org once))]
    (testing "org export is idempotent and keeps the document structure"
      (is (= once twice))
      (is (= (structure (organ/parse-org src)) (structure (organ/parse-org once)))))
    (testing "the export keeps the source's text"
      (is (str/starts-with? once (str "#+TITLE: Round trip\n#+OPTIONS: toc:t num:t\n"
                                      "#+TODO: TODO NEXT | DONE CANX\n#+STARTUP: overview\n"
                                      "#+HTML_HEAD: <a>\n#+HTML_HEAD: <b>\n\n"))
          "the whole preamble is re-emitted in file order")
      (is (str/includes? once "* NEXT [#A] Task :tag:") "custom keyword survives the round trip")
      (is (str/includes? once ":LOGBOOK:\n- Note\n:END:"))
      (is (str/includes? once "#+PROPERTY: header-args :results silent") "body keyword lines are kept")
      (is (str/includes? once "\\begin{equation}\nE = mc^2\n\\end{equation}"))
      (is (str/includes? once "$x$ and \\(y\\)"))
      (is (str/includes? once "- item\n  | a"))
      (is (str/includes? once "  #+FOO: bar\n  :NOTES:\n  hello\n  :END:") "item children are indented")
      (is (str/includes? once "1. one\n\n   continued") "ordered item children sit under the text")
      (is (str/includes? once "#+name: src\n#+BEGIN_SRC clojure -n\n") "affiliated keywords precede blocks")
      (is (str/includes? once "First.\n\nSecond.") "quote paragraphs stay separate")
      (is (str/includes? once "CLOSED: [2025-01-16 09:00]") "CLOSED is an inactive timestamp")
      (is (not (str/includes? once "* 1 ")) "section numbers are not materialised"))
    (testing "body keyword lines are dropped from other formats"
      (is (not (str/includes? (hop/render-ast-as-markdown (organ/parse-org src)) "#+PROPERTY"))))))
(deftest org-roundtrip-lists-test
  (doseq [[label src] [["counter-set cookie" "1. a\n2. b\n5. [@5] e\n6. f\n"]
                       ["two lists split by two blank lines" "- a\n- b\n\n\n- c\n"]
                       ["description list then plain list" "- a :: def\n\n\n- plain\n"]]]
    (testing label
      (let [once (hop/render-ast-as-org (organ/parse-org src))]
        (is (= (structure (organ/parse-org src)) (structure (organ/parse-org once))))
        (is (= once (hop/render-ast-as-org (organ/parse-org once)))))))
  (testing "the cookie restarts the numbering"
    (is (str/includes? (hop/render-ast-as-org (organ/parse-org "1. a\n2. [@5] b\n3. c"))
                       "1. a\n5. [@5] b\n6. c")))
  (testing "html gives a restarted item its value"
    (is (str/includes? (hop/render-ast-as-html (organ/parse-org "1. a\n2. [@5] b"))
                       "<li value=\"5\">b</li>"))))

(deftest org-roundtrip-inline-footnote-test
  (testing "an inline footnote keeps its markup and its links"
    (let [src  "Text[fn::see [[https://x.org][the *site*]] ok] and [fn:n:a /note/].\n"
          once (hop/render-ast-as-org (organ/parse-org src))]
      (is (str/includes? once "[fn::see [[https://x.org][the *site*]] ok]"))
      (is (str/includes? once "[fn:n:a /note/]"))
      (is (= (structure (organ/parse-org src)) (structure (organ/parse-org once)))))))

(deftest org-roundtrip-affiliated-test
  (testing "#+header: and #+results: lines are re-emitted"
    (let [src  (str "#+NAME: blk\n#+HEADER: :var x=1\n#+BEGIN_SRC sh\necho $x\n#+END_SRC\n\n"
                    "#+RESULTS[abc123]: blk\n| 1 |\n")
          once (hop/render-ast-as-org (organ/parse-org src))]
      (is (str/includes? once "#+name: blk\n#+header: :var x=1\n#+BEGIN_SRC sh"))
      (is (str/includes? once "#+results: blk\n| 1"))
      (is (= (structure (organ/parse-org src)) (structure (organ/parse-org once))))
      (is (= once (hop/render-ast-as-org (organ/parse-org once)))))))

(deftest cli-options-test
  (testing "short flags are unique (tools.cli does not check this)"
    (let [shorts (keep first cli-options)]
      (is (= (count shorts) (count (set shorts))) (str "duplicates: " (frequencies shorts)))))
  (testing "-d is an integer duration and -D a keyword set"
    (let [{:keys [options errors]} (cli/parse-opts ["-d" "45" "-D" "CANX, CANCELED" "-z" "Europe/Paris"] cli-options)]
      (is (empty? errors))
      (is (= 45 (:default-duration options)))
      (is (= #{"CANX" "CANCELED"} (:done-keywords options)))
      (is (= (java.time.ZoneId/of "Europe/Paris") (:time-zone options)))))
  (testing "invalid values are rejected by the parser"
    (is (seq (:errors (cli/parse-opts ["-d" "DONE"] cli-options))))
    (is (seq (:errors (cli/parse-opts ["-z" "Mars/Olympus"] cli-options))))
    (is (seq (:errors (cli/parse-opts ["-f" "pdf"] cli-options))))))
