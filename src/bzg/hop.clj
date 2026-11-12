#!/usr/bin/env bb

;; Copyright (c) Bastien Guerry
;; SPDX-License-Identifier: EPL-2.0
;; License-Filename: LICENSES/EPL-2.0.txt

(ns bzg.hop
  "Hush Org Parser - Render and export Org ASTs produced by organ."
  (:require [clojure.string :as str]
            [cheshire.core :as json]
            [clojure.tools.cli :as cli]
            [babashka.process :as process]
            [clojure.edn :as edn]
            [clj-yaml.core :as yaml]
            [bzg.organ :as organ]))

;; Re-bind organ utilities for local use
(def ^:private remove-empty-vals organ/remove-empty-vals)
(def ^:private non-blank? organ/non-blank?)
(defn- escape-block-content
  "organ/escape-block-content, for a block that may be empty (organ then
   leaves :content out of the node)."
  [content]
  (organ/escape-block-content (or content "")))
(def ^:private parse-inline organ/parse-inline)
(def ^:private inline-text organ/inline-text)

;; Rendering constants
(def ^:private min-table-cell-width 3)
(def ^:private max-heading-level 6)
(def ^:private list-indent-width 2)

;; Dynamic vars for rendering
(def ^:dynamic *base-url* nil)
(def ^:dynamic *css-theme* nil)
(def ^:private pico-themes-cdn "https://cdn.jsdelivr.net/gh/bzg/pico-themes@latest/")

;; Inline Node Rendering
;; Render organ's parsed inline nodes to HTML, Markdown or Org strings.

(defn- escape-html
  "HTML-escape text; returns it untouched when nothing needs escaping,
   which is 4x faster on large strings such as base64 data URIs."
  [text]
  (let [s (str text)]
    (if (re-find #"[&<>\"]" s)
      (str/escape s {\& "&amp;" \< "&lt;" \> "&gt;" \" "&quot;"})
      s)))

(defn- upper-name [k] (str/upper-case (name k)))

(defn- repeat-str
  "(str/join (repeat n s)) — empty string when n <= 0."
  [n s]
  (str/join (repeat n s)))

(defn- prefix-lines
  "Prefix each line of content with the given string."
  [prefix content]
  (->> (str/split-lines content)
       (map #(str prefix %))
       (str/join "\n")))

(defn- indent-lines
  "Indent every non-blank line of s by n spaces."
  [n s]
  (let [prefix (repeat-str n " ")]
    (->> (str/split-lines s)
         (map #(if (str/blank? %) % (str prefix %)))
         (str/join "\n"))))

(defn- dedent
  "Strip the common leading whitespace of the non-blank lines of s, as Org
   does for the content of blocks and drawers nested in list items."
  [s]
  (let [lines (str/split-lines s)
        n     (->> lines
                   (remove str/blank?)
                   (map #(count (re-find #"^\s*" %)))
                   (reduce min Integer/MAX_VALUE))]
    (if (or (zero? n) (= n Integer/MAX_VALUE))
      s
      (str/join "\n" (map #(if (str/blank? %) % (subs % n)) lines)))))

(def ^:private image-mime-types
  {"png"  "image/png"
   "jpg"  "image/jpeg"
   "jpeg" "image/jpeg"
   "gif"  "image/gif"
   "svg"  "image/svg+xml"
   "webp" "image/webp"
   "bmp"  "image/bmp"
   "ico"  "image/x-icon"
   "tiff" "image/tiff"
   "tif"  "image/tiff"})

(def ^:private image-extensions (set (keys image-mime-types)))

(defn- get-file-extension
  "Extract lowercase file extension from a path or URL."
  [path]
  (when path
    (-> path str/lower-case (str/split #"[?#]") first (str/split #"\.") last)))

(defn- image-url?
  "Check if a URL points to an image based on file extension."
  [url]
  (when url
    (contains? image-extensions (get-file-extension url))))

(defn- expand-home
  "Expand ~ to user home directory in file paths."
  [path]
  (if (and path (str/starts-with? path "~/"))
    (str (System/getProperty "user.home") (subs path 1))
    path))

(defn- file-to-base64
  "Base64-encode a file's content, or nil if unreadable."
  [filepath]
  (try
    (let [file (java.io.File. (expand-home filepath))]
      (when (.exists file)
        (-> file .toPath java.nio.file.Files/readAllBytes
            (->> (.encodeToString (java.util.Base64/getEncoder))))))
    (catch Exception _ nil)))

(def ^:private image-to-data-uri
  "Convert a local image file to a data URI, or nil if not possible.
   Memoized so repeated references to the same image are read and
   encoded only once per run."
  (memoize
   (fn [filepath]
     (when-let [mime (some-> (get-file-extension filepath) image-mime-types)]
       (when-let [b64 (file-to-base64 filepath)]
         (str "data:" mime ";base64," b64))))))

(defn- build-img-attrs
  "Build an HTML attribute string from a map of attributes."
  [attrs]
  (when (seq attrs)
    (->> attrs
         (keep (fn [[k v]]
                 (when (non-blank? (str v))
                   (str " " (name k) "=\"" (escape-html (str v)) "\""))))
         (str/join))))

(defn- render-image-html
  "Render an <img> tag merging #+attr_html attributes; default-alt is used
   when attrs provide no :alt. A :caption wraps the tag in a <figure>."
  [src default-alt affiliated]
  (let [html-attrs (get-in affiliated [:attr :html] {})
        alt        (or (:alt html-attrs) default-alt)
        img-attrs  (merge {:src src :alt alt} (dissoc html-attrs :alt))
        img-tag    (str "<img" (build-img-attrs img-attrs) ">")
        caption    (:caption affiliated)]
    (if caption
      (str "<figure>" img-tag "<figcaption>"
           (escape-html caption) "</figcaption></figure>")
      img-tag)))

(defn- heading-to-slug
  "URL-safe anchor slug from a heading title (string or inline nodes)."
  [title]
  (let [text (if (string? title) title (or (inline-text title) ""))
        slug (-> text
                 str/lower-case
                 (str/replace #"[^\p{L}\p{N}\s-]" "")
                 str/trim
                 (str/replace #"\s+" "-"))]
    (if (str/blank? slug) "section" slug)))

(defn- prepend-base-url
  "Prepend *base-url* to href when it is a relative path."
  [href]
  (if (and href *base-url*
           (not (str/starts-with? href "/"))
           (not (str/starts-with? href "#"))
           (not (re-find #"^[a-zA-Z][a-zA-Z0-9+.-]*:" href)))
    (str *base-url* href)
    href))

(defn- resolve-link-href
  "Resolve a link node to a URL string. file: links get .org -> .html
   for html/md output; org output keeps the raw link verbatim."
  [{:keys [link-type target url]} fmt]
  (prepend-base-url
   (case link-type
     :file        (if (= fmt :org) url (str/replace target #"\.org$" ".html"))
     (:id
      :custom-id) (str "#" target)
     :heading     (str "#" (heading-to-slug (parse-inline target)))
     :mailto      (str "mailto:" target)
     url)))

;; Inline Node Rendering -- forward declarations
(declare render-inline render-link-node)

(def ^:private inline-markup
  "[open close] delimiters of emphasis and literal inline nodes, per format."
  {:bold      {:html ["<strong>" "</strong>"] :md ["**" "**"] :org ["*" "*"]}
   :italic    {:html ["<em>" "</em>"]         :md ["*" "*"]   :org ["/" "/"]}
   :underline {:html ["<u>" "</u>"]           :md ["_" "_"]   :org ["_" "_"]}
   :strike    {:html ["<del>" "</del>"]       :md ["~~" "~~"] :org ["+" "+"]}
   :code      {:html ["<code>" "</code>"]     :md ["`" "`"]   :org ["~" "~"]}
   :verbatim  {:html ["<code>" "</code>"]     :md ["`" "`"]   :org ["=" "="]}})

(def ^:private latex-fragment-delimiters
  "Delimiters of an inline LaTeX fragment, by the :kind organ parsed."
  {:paren   ["\\(" "\\)"]
   :bracket ["\\[" "\\]"]
   :dollars ["$$" "$$"]
   :dollar  ["$" "$"]})

(defn- wrap-markup
  "Wrap text in the delimiters of inline node type `type` for fmt."
  [type fmt text]
  (let [[open close] (get-in inline-markup [type fmt])]
    (str open text close)))

(defn- render-inline-node
  "Render a single inline node to a string in the given format."
  [node fmt]
  (case (:type node)
    :text      (case fmt :html (escape-html (:value node)) (:value node))
    (:bold :italic :underline :strike)
    (wrap-markup (:type node) fmt (render-inline (:children node) fmt))
    (:code :verbatim)
    (wrap-markup (:type node) fmt
                 (case fmt :html (escape-html (:value node)) (:value node)))
    :link      (render-link-node node fmt nil)
    :footnote-ref
    (let [label (:label node)]
      (case fmt
        :html (str "<sup id=\"fnref-" label "\"><a href=\"#fn-"
                   label "\" class=\"footnote-ref\">" label "</a></sup>")
        :md   (str "[^" label "]")
        (str "[fn:" label "]")))
    :footnote-inline
    (let [label    (:label node)
          children (:children node)]
      (case fmt
        :html (let [display (if (str/blank? label) "*" label)
                    text    (escape-html (or (inline-text children) ""))]
                (str "<sup><a href=\"#\" title=\"" text
                     "\" class=\"footnote-inline\">" display "</a></sup>"))
        :md   (str "[^" (or label "") "]")
        ;; org: keep the definition's markup (emphasis, links)
        (str "[fn:" (or label "") ":" (render-inline children fmt) "]")))
    (:timestamp :macro) (cond-> (:raw node) (= fmt :html) escape-html)
    :latex-fragment
    (let [[open close] (latex-fragment-delimiters (:kind node))
          text         (str open (:value node) close)]
      (case fmt :html (escape-html text) text))
    ;; fallback
    (or (:value node) "")))

(defn- render-inline
  "Render inline nodes (or a fallback string) to a string in fmt."
  [nodes fmt]
  (cond
    (nil? nodes)    ""
    (string? nodes) (case fmt :html (escape-html nodes) nodes)
    (empty? nodes)  ""
    :else           (let [sb (StringBuilder.)]
                      (doseq [node nodes]
                        (.append sb (render-inline-node node fmt)))
                      (.toString sb))))

(defn- extract-single-image-node
  "Return the single inline node when it is an image link, else nil."
  [nodes]
  (when (and (sequential? nodes)
             (= 1 (count nodes))
             (= :link (:type (first nodes))))
    (let [node       (first nodes)
          actual-url (if (#{:http :https} (:link-type node))
                       (:url node) (:target node))]
      (when (image-url? actual-url) node))))

(defn- html-img
  "Plain <img> tag; alt defaults to src."
  ([src] (html-img src src))
  ([src alt]
   (str "<img src=\"" (escape-html src) "\" alt=\"" (escape-html alt) "\">")))

(defn- html-linked-img
  "An image (used as link description) wrapped in a link to href."
  [href img-src]
  (str "<a href=\"" (escape-html href) "\">" (html-img img-src) "</a>"))

(defn- render-link-node
  "Render a :link inline node; affiliated keywords (e.g. #+CAPTION)
   enable enhanced image rendering."
  [node fmt affiliated]
  (let [link-type     (:link-type node)
        target        (:target node)
        url           (:url node)
        children      (:children node)
        has-desc      (seq children)
        desc-text     (when has-desc (inline-text children))
        actual-url    (if (#{:http :https} link-type) url (or target url))
        is-local-file (= link-type :file)
        is-remote     (#{:http :https} link-type)
        url-is-image  (image-url? actual-url)
        desc-is-image (and desc-text (image-url? desc-text))]
    (cond
      ;; Local file images: html/md get a base64 data URI, falling back to
      ;; the raw path when the file is unreadable. Org output falls through
      ;; to the generic branch and keeps the link verbatim.
      (and is-local-file url-is-image (#{:html :md} fmt))
      (let [path (or target url)
            src  (or (image-to-data-uri path) path)]
        (cond
          (= fmt :md) (str "![" (or desc-text target) "](" src ")")
          affiliated    (render-image-html src (or desc-text target) affiliated)
          desc-is-image (html-linked-img src desc-text)
          :else         (html-img src (or desc-text target))))

      ;; HTML format
      (= fmt :html)
      (cond
        (and is-remote url-is-image (not has-desc))
        (if affiliated
          (render-image-html url url affiliated)
          (html-img url))

        (and is-remote url-is-image has-desc (not desc-is-image))
        (if affiliated
          (render-image-html url desc-text affiliated)
          (html-img url desc-text))

        (and is-remote desc-is-image)
        (html-linked-img url desc-text)

        :else
        (let [href    (escape-html (resolve-link-href node :html))
              display (if has-desc
                        (render-inline children :html)
                        (escape-html (case link-type
                                       (:heading :custom-id) target
                                       url)))]
          (str "<a href=\"" href "\">" display "</a>")))

      ;; Markdown format
      (= fmt :md)
      (cond
        (and is-remote url-is-image (not has-desc))
        (str "![" url "](" url ")")

        (and is-remote desc-is-image)
        (str "[![" desc-text "](" desc-text ")](" url ")")

        :else
        (let [href    (resolve-link-href node :md)
              display (if has-desc
                        (render-inline children :md)
                        (case link-type
                          (:heading :custom-id) target
                          url))]
          (str "[" display "](" href ")")))

      ;; Org format
      :else
      (let [href    (-> (resolve-link-href node :org) ; as org-link-escape
                        (str/replace #"(\\*)(?=[\[\]]|\z)" "$1$1")
                        (str/replace #"[\[\]]" "\\\\$0"))
            display (when has-desc (render-inline children :org))]
        (if display
          (str "[[" href "][" display "]]")
          (str "[[" href "]]"))))))

;; AST Content Rendering
(defn- render-content-in-node [node render-format]
  (let [fmt              #(render-inline % render-format)
        render-children* #(mapv (fn [c] (render-content-in-node c render-format)) %)
        result           (case (:type node)
                           :document     (-> node (update :title #(when % (fmt %))) (update :children render-children*))
                           :section      (-> node (update :title #(when % (fmt %))) (update :children render-children*))
                           :paragraph    (update node :content fmt)
                           :list         (update node :items #(mapv (fn [i] (render-content-in-node i render-format)) %))
                           :list-item    (let [node (update node :children render-children*)]
                                           (if (:term node)
                                             (-> node (update :term fmt) (update :definition fmt))
                                             (update node :content fmt)))
                           :table        (update node :rows #(mapv (fn [row] (mapv fmt row)) %))
                           :quote-block  (update node :children render-children*)
                           :footnote-def (update node :content fmt)
                           :block        (if (#{:src :example :export} (:block-type node)) node
                                             (update node :content
                                                     #(render-inline (parse-inline %) render-format)))
                           node)]
    (remove-empty-vals result)))

;; Unified Renderer
(def ^:private html-styles
  "body { font-family: sans-serif; line-height: 1.6; margin: 2em auto; max-width: 800px; padding: 0 1em; }
h1, h2, h3, h4, h5, h6 { line-height: 1.2; }
header { margin-bottom: 1.5em; }
header .document-title { margin-bottom: 0.2em; padding-bottom: 0.2em; border-bottom: 3px solid; border-color: #ccc; }
.subtitle { color: #555; font-size: 1.2em; margin-top: 0; }
pre { background-color: #f8f8f8; border: 1px solid #ddd; border-radius: 4px; padding: 1em; overflow-x: auto; }
code { font-family: monospace, monospace; }
img { max-width: 100%; }
figure { margin: 1em 0; }
figure img { display: block; }
figcaption { font-size: 0.9em; color: #555; margin-top: 0.5em; font-style: italic; }
pre code { background-color: transparent; border: none; padding: 0; }
blockquote { border-left: 4px solid #eee; margin-left: 0; padding-left: 1em; color: #555; }
table { border-collapse: collapse; width: 100%; margin-bottom: 1em; }
th, td { border: 1px solid #ddd; padding: 0.5em; text-align: left; }
th { background-color: #f2f2f2; font-weight: bold; }
ul, ol { padding-left: 2em; }
li > p { margin-top: 0.5em; }
.todo { font-weight: bold; color: #c00; }
.done { font-weight: bold; color: #0a0; }
.priority { color: #c60; }
.tags { color: #666; font-size: 0.9em; }
.footnotes { margin-top: 2em; padding-top: 1em; border-top: 1px solid #ccc; font-size: 0.9em; }
.footnote { margin: 0.5em 0; }
.footnote-ref, .footnote-inline { text-decoration: none; }
.footnote-inline { cursor: help; border-bottom: 1px dotted #666; }
.warning { background: #fff3cd; padding: 0.5em; border-left: 4px solid #ffc107; margin: 1em 0; }
.note { background: #d1ecf1; padding: 0.5em; border-left: 4px solid #17a2b8; margin: 1em 0; }
.planning { color: #666; font-size: 0.9em; margin: 0.2em 0 0.5em 0; }
.planning-keyword { font-weight: bold; }")

(def ^:private hljs-cdn "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.11.2")

(defn- html-template [title content]
  (let [has-code (str/includes? content "<pre><code")]
    (str "<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n"
         "  <meta charset=\"UTF-8\">\n"
         "  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">\n"
         "  <title>" (escape-html title) "</title>\n"
         (when *css-theme*
           (if-let [url (:link *css-theme*)]
             (str "  <link rel=\"stylesheet\" href=\"" url "\">\n")
             (str "  <style>\n" (:inline *css-theme*) "\n  </style>\n")))
         (when has-code
           (str "  <link rel=\"stylesheet\" href=\"" hljs-cdn "/styles/default.min.css\">\n"))
         "  <style>\n" html-styles "\n  </style>\n"
         "</head>\n<body>\n" content
         (when has-code
           (str "\n<script src=\"" hljs-cdn "/highlight.min.js\"></script>"
                "\n<script>hljs.highlightAll();</script>"))
         "\n</body>\n</html>")))

(defn- render-table
  "Render table rows as an aligned pipe table (org or markdown)."
  [rows has-header fmt]
  (if (empty? rows) ""
      (let [;; Find max number of columns across all rows
            max-cols      (reduce max 0 (map count rows))
            ;; Pad helper for ragged rows
            pad-row       (fn [row] (let [missing (- max-cols (count row))]
                                      (if (pos? missing)
                                        (into (vec row) (repeat missing []))
                                        (vec row))))
            ;; Render all cells to strings, then compute widths from rendered text
            rendered-rows (mapv (fn [row] (mapv #(render-inline % fmt) (pad-row row))) rows)
            col-widths    (when (and (seq rendered-rows) (pos? max-cols))
                            (apply mapv (fn [& cells]
                                          (apply max min-table-cell-width (map count cells)))
                                   rendered-rows))
            pad-cell      (fn [cell width] (str cell (repeat-str (- width (count cell)) " ")))
            format-row    (fn [row]
                            (str "| " (str/join " | " (map-indexed #(pad-cell %2 (nth col-widths %1 min-table-cell-width)) row)) " |"))
            separator     (when (seq col-widths)
                            (str "|-" (str/join (if (= fmt :org) "-+-" "-|-") (map #(repeat-str % "-") col-widths)) "-|"))]
        (if (nil? col-widths)
          ""
          (if has-header
            (str (format-row (first rendered-rows)) "\n" separator "\n"
                 (str/join "\n" (map format-row (rest rendered-rows))))
            (str/join "\n" (map format-row rendered-rows)))))))

(declare ^:private render-node)

(defn- render-properties-org
  "Render properties as org :PROPERTIES: drawer."
  [properties]
  (when (seq properties)
    (->> properties
         (map (fn [[k v]] (str ":" (upper-name k) ": " v)))
         (str/join "\n")
         (format ":PROPERTIES:\n%s\n:END:"))))

(defn- render-list-item
  "Render one list item (org or markdown). Children are rendered at column
   0 and indented under the marker, so nesting composes: a nested list is
   indented again by its own parent. number is the item's number in an
   ordered list (see list-numbers); an item that restarts the numbering
   gets back its [@N] cookie in org."
  [item number ordered fmt]
  (let [children         (:children item)
        marker           (if ordered (str number ". ") "- ")
        cookie           (when (and ordered (= fmt :org) (:counter item))
                           (str "[@" (:counter item) "] "))
        has-block-child? (and (= fmt :md)
                              (some #(#{:table :src-block :quote-block :block} (:type %)) children))
        ;; A paragraph child needs a blank line before it, or it is read
        ;; back as a continuation of the preceding text.
        separator        (fn [child]
                           (if (or has-block-child? (= :paragraph (:type child))) "\n\n" "\n"))
        children-str     (when (seq children)
                           (indent-lines (if ordered (count marker) list-indent-width)
                                         (str/join (map #(str (separator %) (render-node % fmt))
                                                        children))))]
    (if (and (:term item) (= fmt :md))
      (str (render-inline (:term item) :md) "\n"
           ":   " (render-inline (or (:definition item) []) :md)
           children-str)
      (let [content (if (:term item)
                      (str (render-inline (:term item) fmt) " :: "
                           (render-inline (or (:definition item) []) fmt))
                      (render-inline (:content item) fmt))]
        (str marker cookie content children-str)))))

(defn- render-children
  "Render a sequence of child nodes with smart spacing.
   Sections handle their own spacing via :blank-lines-before."
  [children fmt]
  (let [separator (fn [prev child]
                    (cond
                      (or (= (:type child) :section) (= fmt :html)) "\n"
                      ;; Org reads two lists separated by a single blank
                      ;; line as one list: two blank lines end a list
                      (and (= fmt :org) (= :list (:type prev) (:type child))) "\n\n\n"
                      :else "\n\n"))
        rendered  (keep (fn [child]
                          (let [r (render-node child fmt)]
                            (when-not (str/blank? r) [child r])))
                        children)]
    (str/join (map (fn [[prev _] [child r]] (if prev (str (separator prev child) r) r))
                   (cons nil rendered)
                   rendered))))

(defn- meta-string
  "A document :meta value as one string. A repeated #+KEY: line gives a
   vector, which Org joins with a space."
  [ast k]
  (let [v (get-in ast [:meta k])]
    (if (vector? v) (str/join " " v) v)))

;; #+OPTIONS parsing
(defn- parse-options-string
  "Parse an #+OPTIONS: value like 'toc:t H:2 num:t' into a map
   ('t' -> true, 'nil' -> false, integers parsed, else string)."
  [s]
  (when (and s (not (str/blank? s)))
    (into {}
          (for [[_ k v] (re-seq #"(\S+):(\S+)" s)]
            [(keyword (str/lower-case k))
             (case v
               "t"   true
               "nil" false
               (try (Integer/parseInt v) (catch Exception _ v)))]))))

(defn- collect-toc-entries
  "Recursively collect TOC entries from section nodes up to max-depth.
   Returns a flat list of {:level :title :section-number :id} maps."
  [children max-depth]
  (reduce
   (fn [entries child]
     (if (and (= (:type child) :section)
              (<= (:level child) max-depth))
       (let [entry {:level          (:level child)
                    :title          (:title child)
                    :section-number (:section-number child)
                    :id             (or (:section-id child)
                                        (get-in child [:properties :custom_id])
                                        (:slug child))}]
         (into (conj entries entry)
               (collect-toc-entries (:children child) max-depth)))
       entries))
   []
   children))

(defn- normalize-toc-levels
  "Clamp level jumps to +1 (a '*' heading followed by '***') so TOC
   nesting stays well-formed."
  [entries]
  (loop [[e & more] entries prev 0 acc []]
    (if (nil? e)
      acc
      (let [lvl (min (:level e) (inc prev))]
        (recur more lvl (conj acc (assoc e :level lvl)))))))

(defn- render-toc
  "Render a table of contents from TOC entries in the given format."
  [entries fmt]
  (when-let [entries (seq (normalize-toc-levels entries))]
    (case fmt
      :html (let [;; Markup emitted before an entry, given the previous entry's level:
                  ;; open sublists when going deeper, close them when coming back up,
                  ;; and close the previous item whenever one is open at this level.
                  before (fn [prev level]
                           (cond
                             (> level prev) (repeat-str (- level prev) "<ul>\n")
                             (= level prev) "</li>\n"
                             :else          (str (repeat-str (- prev level) "</li>\n</ul>\n")
                                                 "</li>\n")))
                  link   (fn [{:keys [title section-number id]}]
                           (str "<li><a href=\"#" (escape-html id) "\">"
                                (when section-number (str section-number " "))
                                (render-inline title :html) "</a>"))
                  levels (cons 0 (map :level entries))
                  last-level (:level (last entries))]
              (str "<nav id=\"table-of-contents\">\n<h2>Table of Contents</h2>\n"
                   (str/join (map (fn [prev e] (str (before prev (:level e)) (link e)))
                                  levels entries))
                   (repeat-str last-level "</li>\n</ul>\n")
                   "</nav>"))
      ;; markdown (org output never materialises the TOC)
      (->> entries
           (map (fn [{:keys [level title section-number]}]
                  (str (repeat-str (* 2 (dec level)) " ") "- ["
                       (when section-number (str section-number " "))
                       (render-inline title :md) "](#" (heading-to-slug title) ")")))
           (str/join "\n")))))

(defn- org-keyword-line
  "#+KEY: value line; a blank value gives a bare #+KEY:, as in the source."
  [k v]
  (str "#+" (upper-name k) ":" (when (non-blank? (str v)) (str " " v))))

(defn- render-meta-org
  "Re-emit the document preamble as #+KEY: lines. organ keeps every
   preamble keyword in :meta, in file order; a repeated keyword holds a
   vector of values."
  [meta]
  (->> meta
       (mapcat (fn [[k v]] (map #(org-keyword-line k %) (if (vector? v) v [v]))))
       (str/join "\n")))

(defn- render-document [node fmt]
  (let [{toc? :toc h :h} (:parsed-options node)
        ;; toc can be true (use H:) or an integer depth; org output keeps
        ;; the #+OPTIONS line instead of materialising the TOC
        toc-depth   (cond (integer? toc?) toc? (integer? h) h :else 6)
        toc-str     (when (and toc? (not= fmt :org))
                      (render-toc (collect-toc-entries (:children node) toc-depth) fmt))]
    (case fmt
      :html (let [footnotes      (filter #(= (:type %) :footnote-def) (:children node))
                  title-html     (when-let [t (:title node)]
                                   (str "<header>\n<h1 class=\"document-title\">" (render-inline t :html) "</h1>\n"
                                        (when-let [st (meta-string node :subtitle)]
                                          (str "<p class=\"subtitle\">" (render-inline (parse-inline st) :html) "</p>\n"))
                                        "</header>\n"))
                  toc-html       (when toc-str (str toc-str "\n"))
                  main-content   (str title-html toc-html
                                      (->> (:children node)
                                           (remove #(= (:type %) :footnote-def))
                                           (map #(render-node % fmt))
                                           (str/join "\n")))
                  footnotes-html (when (seq footnotes)
                                   (str "<aside class=\"footnotes\">\n"
                                        (str/join "\n" (map #(render-node % fmt) footnotes))
                                        "\n</aside>"))]
              (html-template (or (inline-text (:title node)) "Untitled Document")
                             (str main-content (when footnotes-html (str "\n" footnotes-html)))))
      ;; Org output is source-faithful: the #+OPTIONS line carries the TOC
      ;; and numbering intent, so neither is materialised in the text.
      :org  (let [meta-str (render-meta-org (:meta node))]
              (str (when (seq meta-str) (str meta-str "\n\n"))
                   (render-children (:children node) fmt)))
      ;; default: markdown
      (let [title (if-let [t (:title node)] (str "# " (render-inline t :md) "\n\n") "")]
        (str title
             (when toc-str (str toc-str "\n\n"))
             (render-children (:children node) fmt))))))

(def ^:private planning-order
  "Canonical order for planning keywords."
  [:closed :deadline :scheduled])

(defn- planning-repeat
  "Repeater of a planning keyword (:scheduled -> :scheduled-repeat)."
  [planning kw]
  (get planning (keyword (str (name kw) "-repeat"))))

(defn- planning-items
  "Return [[kw iso] ...] pairs in canonical order, filtering nil values."
  [planning]
  (keep (fn [kw] (when-let [iso (get planning kw)] [kw iso])) planning-order))

(defn- iso->org-timestamp
  "Org timestamp for an ISO value (2025-01-15T10:30, or a start/end range),
   with an optional repeater cookie: <2025-01-15 10:30-11:00 +1w>. A range
   with a time on both sides shares one timestamp; otherwise it is written
   as <start>--<end>. inactive? gives [brackets], as Org writes CLOSED."
  ([iso repeater] (iso->org-timestamp iso repeater false))
  ([iso repeater inactive?]
   (let [[open close] (if inactive? ["[" "]"] ["<" ">"])
         rep-str      (if repeater (str " " repeater) "")
         single       (fn [s] (str open (str/replace s "T" " ") rep-str close))
         timed?       #(str/includes? % "T")]
     (if (str/includes? iso "/")
       (let [[start end] (str/split iso #"/")]
         (if (and (timed? start) (timed? end))
           (str open (str/replace start "T" " ") "-" (second (str/split end #"T")) rep-str close)
           (str (single start) "--" (single end))))
       (single iso)))))

(defn- render-planning
  "Render planning info (CLOSED/SCHEDULED/DEADLINE) in the given format."
  [planning fmt]
  (when (seq planning)
    (let [items (planning-items planning)]
      (when (seq items)
        (case fmt
          :html (str "<div class=\"planning\">"
                     (str/join " "
                               (map (fn [[kw iso]]
                                      (let [rep      (planning-repeat planning kw)
                                            datetime (first (str/split iso #"/"))
                                            display  (cond-> (str/replace iso "/" "–")
                                                       rep (str " " rep))]
                                        (str "<span class=\"planning-keyword\">"
                                             (str/upper-case (name kw))
                                             ":</span> <time datetime=\"" datetime "\">" display "</time>")))
                                    items))
                     "</div>")
          :org  (str/join " "
                          (map (fn [[kw iso]]
                                 (str (str/upper-case (name kw)) ": "
                                      (iso->org-timestamp iso (planning-repeat planning kw)
                                                          (= kw :closed))))
                               items))
          ;; markdown
          (str/join " "
                    (map (fn [[kw iso]]
                           (let [rep     (planning-repeat planning kw)
                                 display (if rep (str iso " " rep) iso)]
                             (str "**" (str/upper-case (name kw)) ":** " display)))
                         items)))))))

(defn- render-section [node fmt]
  (let [planning-str (render-planning (:planning node) fmt)]
    (case fmt
      :html (let [lvl           (min (:level node) max-heading-level)
                  tag           (str "h" lvl)
                  section-id    (or (:section-id node)
                                    (get-in node [:properties :custom_id])
                                    (heading-to-slug (:title node)))
                  sec-num-html  (when-let [sn (:section-number node)]
                                  (str "<span class=\"section-number\">" sn "</span> "))
                  todo-html     (when (:todo node)
                                  (str "<span class=\"todo " (str/lower-case (name (:todo node))) "\">"
                                       (name (:todo node)) "</span> "))
                  priority-html (when (:priority node)
                                  (str "<span class=\"priority\">[#" (:priority node) "]</span> "))
                  tags-html     (when (seq (:tags node))
                                  (str " <span class=\"tags\">:" (str/join ":" (:tags node)) ":</span>"))]
              (str "<section id=\"" (escape-html section-id) "\">\n<" tag ">"
                   sec-num-html todo-html priority-html (render-inline (:title node) :html) tags-html
                   "</" tag ">\n"
                   (when planning-str (str planning-str "\n"))
                   (render-children (:children node) fmt) "\n</section>"))
      :org  (let [blanks-before      (repeat-str (or (:blank-lines-before node) 0) "\n")
                  blanks-after-title (repeat-str (or (:blank-lines-after-title node) 0) "\n")
                  stars              (repeat-str (:level node) "*")
                  todo-str           (when (:todo node) (str (name (:todo node)) " "))
                  priority-str       (when (:priority node) (str "[#" (:priority node) "] "))
                  tags-str           (when (seq (:tags node)) (str " :" (str/join ":" (:tags node)) ":"))
                  props-str          (render-properties-org (:properties node))
                  children-str       (render-children (:children node) fmt)]
              (str blanks-before
                   stars " " todo-str priority-str (render-inline (:title node) :org) tags-str
                   (when planning-str (str "\n" planning-str))
                   (when props-str (str "\n" props-str))
                   blanks-after-title
                   (when (seq children-str) (str "\n" children-str))))
      ;; default: markdown
      (let [blanks-before      (repeat-str (or (:blank-lines-before node) 0) "\n")
            blanks-after-title (repeat-str (or (:blank-lines-after-title node) 0) "\n")
            todo-str           (when (:todo node) (str "**" (name (:todo node)) "** "))
            priority-str       (when (:priority node) (str "[#" (:priority node) "] "))
            sec-num-str        (when-let [sn (:section-number node)] (str sn " "))
            tags-str           (when (seq (:tags node)) (str " `:" (str/join ":" (:tags node)) ":`"))
            heading            (str (repeat-str (min (:level node) max-heading-level) "#") " "
                                    todo-str priority-str sec-num-str
                                    (render-inline (:title node) :md)
                                    tags-str)]
        (str blanks-before heading "\n"
             (when planning-str (str planning-str "\n"))
             blanks-after-title (render-children (:children node) fmt))))))

(defn- render-paragraph [node fmt]
  (let [content (:content node)]
    (case fmt
      :html (if-let [img-node (and (:affiliated node) (extract-single-image-node content))]
              (render-link-node img-node :html (:affiliated node))
              (str "<p>" (render-inline content :html) "</p>"))
      ;; org: affiliated keywords are re-emitted by render-node
      (render-inline content fmt))))

(defn- list-numbers
  "The number of each item of an ordered list: one more than the previous
   item, unless the item's [@N] cookie (organ's :counter) restarts the
   count, as in Org."
  [items]
  (rest (reductions (fn [n item] (or (:counter item) (inc n))) 0 items)))

(defn- render-list-node [node fmt]
  (case fmt
    :html (let [tag        (cond (:description node) "dl"
                                 (:ordered node)     "ol"
                                 :else               "ul")
                items-html (->> (:items node)
                                (map #(render-node % fmt))
                                (str/join "\n"))]
            (str "<" tag ">\n" items-html "\n</" tag ">"))
    (let [sep (if (and (= fmt :md) (:description node)) "\n\n" "\n")]
      (->> (map (fn [item n] (render-list-item item n (:ordered node) fmt))
                (:items node) (list-numbers (:items node)))
           (str/join sep)))))

(defn- render-list-item-node [node fmt]
  (if (= fmt :html)
    (let [children-html (when (seq (:children node))
                          (str "\n" (str/join "\n" (map #(render-node % fmt) (:children node)))))]
      (if (:term node)
        (str "<dt>" (render-inline (:term node) :html) "</dt>\n<dd>"
             (render-inline (:definition node) :html) children-html "</dd>")
        (str "<li" (when-let [n (:counter node)] (str " value=\"" n "\"")) ">"
             (render-inline (:content node) :html) children-html "</li>")))
    (render-list-item node 1 false fmt)))

(defn- render-table-node [node fmt]
  (let [rows       (:rows node)
        has-header (:has-header node)]
    (if (= fmt :html)
      (let [cell (fn [tag content] (str "<" tag ">" (render-inline content :html) "</" tag ">"))]
        (if (empty? rows) ""
            (str "<table>\n"
                 (when has-header
                   (str "  <thead>\n    <tr>\n      "
                        (str/join (map #(cell "th" %) (first rows)))
                        "\n    </tr>\n  </thead>\n"))
                 "  <tbody>\n"
                 (str/join "\n"
                           (map (fn [row]
                                  (str "    <tr>\n      "
                                       (str/join (map #(cell "td" %) row))
                                       "\n    </tr>"))
                                (if has-header (rest rows) rows)))
                 "\n  </tbody>\n</table>")))
      (render-table rows has-header fmt))))

(defn- render-src-block [node fmt]
  (case fmt
    :html (str "<pre><code"
               (when (non-blank? (:language node))
                 (str " class=\"language-" (escape-html (:language node)) "\""))
               ">" (escape-html (:content node)) "</code></pre>")
    :org  (let [;; organ's :args is the whole header line (language included)
                header (if (non-blank? (:args node)) (:args node) (:language node))]
            (str "#+BEGIN_SRC"
                 (when (non-blank? header) (str " " header))
                 "\n" (escape-block-content (:content node)) "\n#+END_SRC"))
    (str "```" (or (:language node) "") "\n" (:content node) "\n```")))

(defn- render-quote-block [node fmt]
  (case fmt
    :html (str "<blockquote>\n"
               (str/join "\n" (map #(render-node % fmt) (:children node)))
               "\n</blockquote>")
    :org  (str "#+BEGIN_QUOTE\n"
               (escape-block-content (render-children (:children node) :org))
               "\n#+END_QUOTE")
    (->> (str/split-lines (render-children (:children node) :md))
         (map #(if (str/blank? %) ">" (str "> " %)))
         (str/join "\n"))))

(defn- render-comment-node [node fmt]
  (if (= fmt :org)
    (prefix-lines "# " (:content node))
    (str "<!-- " (escape-html (:content node)) " -->")))

(defn- render-fixed-width [node fmt]
  (case fmt
    :html (str "<pre>" (escape-html (:content node)) "</pre>")
    :org  (prefix-lines ": " (:content node))
    (str "```\n" (:content node) "\n```")))

(defn- render-footnote-def-node [node fmt]
  (case fmt
    :html (let [label (escape-html (:label node))]
            (str "<div class=\"footnote\" id=\"fn-" label "\">"
                 "<a href=\"#fnref-" label "\" class=\"footnote-backref\"><sup>" label "</sup></a> "
                 (render-inline (:content node) :html) "</div>"))
    :org  (str "[fn:" (:label node) "] " (render-inline (:content node) :org))
    (str "[^" (:label node) "]: " (render-inline (:content node) :md))))

(defn- render-generic-block [node fmt]
  (let [block-type  (:block-type node)
        export-type (:args node)]
    (case fmt
      :html (case block-type
              :warning (str "<div class=\"warning\">" (render-inline (parse-inline (:content node)) :html) "</div>")
              :note    (str "<div class=\"note\">" (render-inline (parse-inline (:content node)) :html) "</div>")
              :example (str "<pre>" (escape-html (:content node)) "</pre>")
              :export  (if (= export-type "html") (:content node) "")
              (str "<div class=\"block-" (name block-type) "\">"
                   "<pre>" (escape-html (:content node)) "</pre></div>"))
      :org  (case block-type
              :export (if (= export-type "org")
                        (:content node)
                        (str "#+BEGIN_EXPORT"
                             (when (non-blank? export-type) (str " " export-type))
                             "\n" (escape-block-content (:content node)) "\n#+END_EXPORT"))
              (let [args            (:args node)
                    escaped-content (escape-block-content (:content node))]
                (str "#+BEGIN_" (upper-name block-type)
                     (when (non-blank? args) (str " " args))
                     "\n" escaped-content "\n#+END_" (upper-name block-type))))
      ;; markdown
      (case block-type
        (:warning :note) (str "> **" (str/capitalize (name block-type)) "**\n"
                              (prefix-lines "> " (:content node)))
        :export          (if (= export-type "markdown") (:content node) "")
        :example         (str "```\n" (:content node) "\n```")
        (str "```" (name block-type) "\n" (:content node) "\n```")))))

(defn- render-html-line [node fmt]
  (case fmt
    :html (:content node)
    :org  (prefix-lines "#+html: " (:content node))
    ""))

(defn- render-latex-line [node fmt]
  (case fmt
    :org (prefix-lines "#+latex: " (:content node))
    ""))

(defn- render-affiliated-org
  "Re-emit a node's affiliated keywords (#+caption:, #+name:, #+header:,
   #+results:, #+attr_*:) as the lines that precede it in Org."
  [affiliated]
  (when affiliated
    (->> (concat
          (when-let [caption (:caption affiliated)] [(str "#+caption: " caption)])
          (when-let [aff-name (:name affiliated)] [(str "#+name: " aff-name)])
          (for [k [:header :results] :let [v (get affiliated k)] :when v]
            (str "#+" (name k) ": " v))
          (for [[attr-type attrs] (:attr affiliated)
                :when             (seq attrs)]
            (str "#+attr_" (name attr-type) ": "
                 (str/join " " (for [[k v] attrs]
                                 (if (str/includes? v " ")
                                   (str ":" (name k) " \"" v "\"")
                                   (str ":" (name k) " " v)))))))
         (str/join "\n")
         not-empty)))

(defn- render-node
  "Render a block-level node in fmt (:html, :md or :org). Returns \"\" for
   nodes a format does not show."
  [node fmt]
  (let [rendered
        (case (:type node)
          :document        (render-document node fmt)
          :section         (render-section node fmt)
          :paragraph       (render-paragraph node fmt)
          :list            (render-list-node node fmt)
          :list-item       (render-list-item-node node fmt)
          :table           (render-table-node node fmt)
          :src-block       (render-src-block node fmt)
          :quote-block     (render-quote-block node fmt)
          :property-drawer (if (= fmt :org) (render-properties-org (:properties node)) "")
          :comment         (render-comment-node node fmt)
          :fixed-width     (render-fixed-width node fmt)
          :footnote-def    (render-footnote-def-node node fmt)
          :block           (render-generic-block node fmt)
          :html-line       (render-html-line node fmt)
          :latex-line      (render-latex-line node fmt)
          ;; A #+KEY: value line outside the preamble (organ 3.4+): org only.
          :keyword         (if (= fmt :org) (org-keyword-line (:key node) (:value node)) "")
          ;; Drawers (:LOGBOOK: etc.) are kept in org output only, as Org's own
          ;; exporters drop them by default.
          :drawer          (if (= fmt :org)
                             (str ":" (str/upper-case (:drawer-name node)) ":\n"
                                  (when (non-blank? (:content node))
                                    (str (dedent (:content node)) "\n"))
                                  ":END:")
                             "")
          ;; LaTeX environments are emitted verbatim (\begin..\end lines included),
          ;; as Org does, for MathJax-style rendering downstream.
          :latex-environment (if (= fmt :html)
                               (str "<div class=\"latex-environment\">"
                                    (escape-html (:content node)) "</div>")
                               (:content node))
          ;; Default case for warnings or unknown types
          (if (:warning node)
            (str "<!-- Warning at line " (:error-line node) ": " (:warning node) " -->")
            ""))
        ;; Org output keeps the affiliated keywords of any element
        attr-lines (when (= fmt :org) (render-affiliated-org (:affiliated node)))]
    (if (and attr-lines (not (str/blank? rendered)))
      (str attr-lines "\n" rendered)
      rendered)))

;; Section ID assignment (deduplication)
(defn- assign-section-ids
  "Assign a unique :section-id to each :section node (custom_id when
   available, else the heading slug; duplicates get -1, -2, ... suffixes)."
  [ast]
  (let [used (atom {})]
    (letfn [(unique-id [base]
              (let [n (get @used base 0)]
                (swap! used assoc base (inc n))
                (if (zero? n) base (str base "-" n))))
            (walk [node]
              (case (:type node)
                :section                 (let [slug (heading-to-slug (:title node))
                                               base (or (get-in node [:properties :custom_id]) slug)]
                                           (-> node
                                               (assoc :section-id (unique-id base))
                                               (assoc :slug slug)
                                               (update :children #(mapv walk %))))
                (:document :quote-block) (update node :children #(mapv walk %))
                :list                    (update node :items #(mapv walk %))
                :list-item               (update node :children #(mapv walk %))
                node))]
      (walk ast))))

;; Section numbering
(defn- number-sections
  "Annotate each :section node with a :section-number string ('1', '1.1',
   ...). Only applied when the document has num:t in #+OPTIONS."
  [ast]
  (if-not (get-in ast [:parsed-options :num])
    ast
      (letfn [(number-children [children counters]
                (loop [[child & more] children
                       counters       counters
                       result         []]
                  (if (nil? child)
                    result
                    (if (= (:type child) :section)
                      (let [level          (:level child)
                            ;; Increment counter at this level, reset deeper levels
                            incremented    (update counters level (fnil inc 0))
                            updated        (apply dissoc incremented (filter #(> % level) (keys incremented)))
                            sec-num        (str/join "." (map #(get updated % 1)
                                                              (range 1 (inc level))))
                            numbered-kids  (number-children (:children child) updated)
                            numbered-child (assoc child
                                                  :section-number sec-num
                                                  :children numbered-kids)]
                        (recur more updated (conj result numbered-child)))
                      (recur more counters (conj result child))))))]
        (assoc ast :children (number-children (:children ast) {})))))

;; ICS Export
(defn- ics-fold-line
  "Fold a content line per RFC 5545 (max 75 octets per line)."
  [^String line]
  (let [bytes (.getBytes line "UTF-8")]
    (if (<= (alength bytes) 75)
      line
      (let [sb (StringBuilder.)]
        (loop [offset 0 first? true]
          (let [max-bytes (if first? 75 74)
                remaining (- (alength bytes) offset)]
            (if (<= remaining 0)
              (.toString sb)
              (let [;; Find char boundary at or before max-bytes from offset
                    end   (min (+ offset max-bytes) (alength bytes))
                    ;; Back up if we landed on a UTF-8 continuation byte (10xxxxxx).
                    ;; e == (alength bytes) is a valid cut point (end of buffer).
                    end   (loop [e end]
                            (if (or (<= e offset)
                                    (>= e (alength bytes))
                                    (not= 0x80 (bit-and (aget bytes e) 0xC0)))
                              e
                              (recur (dec e))))
                    chunk (String. bytes offset (- end offset) "UTF-8")]
                (when-not first? (.append sb "\r\n "))
                (.append sb chunk)
                (recur end false)))))))))

(defn- ics-escape
  "Escape text for ICS property values per RFC 5545."
  [s]
  (str/escape s {\\ "\\\\" \, "\\," \; "\\;" \newline "\\n"}))

(def ^:private utc-dtf
  "ICS UTC datetime formatter (20250115T103000Z)."
  (-> (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd'T'HHmmss'Z'")
      (.withZone (java.time.ZoneId/of "UTC"))))

(defn- iso-to-ics-datetime
  "Convert an ISO timestamp (2025-01-15T10:30) to a map of ICS datetimes:
   {:dtstart \"20250115T103000\"}, plus :dtend for intervals (start/end).
   Date-only input stays in VALUE=DATE form (20250115)."
  [iso]
  (let [to-ics (fn [s]
                 (if (str/includes? s "T")
                   (let [[d t]  (str/split s #"T")
                         date   (str/replace d "-" "")
                         digits (str/replace t ":" "")
                         time   (cond-> digits (= (count digits) 4) (str "00"))]
                     (str date "T" time))
                   (str/replace s "-" "")))]
    (if (str/includes? iso "/")
      (let [[start end] (str/split iso #"/")]
        {:dtstart (to-ics start) :dtend (to-ics end)})
      {:dtstart (to-ics iso)})))

(defn- done-item?
  "True if the entry's heading is in a DONE state. Custom keywords are
   parsed by organ from #+TODO:-style directives in the file and from the
   :todo-keywords parse option (fed by -D)."
  [{:keys [todo]} done-keywords]
  (and todo (contains? done-keywords (name todo))))

(defn- prune-done-sections
  "Remove every section subtree (heading + descendants) whose heading is
   in a DONE state."
  [node done-keywords]
  (cond
    (and (map? node) (= :section (:type node))
         (done-item? node done-keywords))
    nil

    (and (map? node) (vector? (:children node)))
    (update node :children #(->> % (keep (fn [c] (prune-done-sections c done-keywords))) vec))

    :else node))

(defn- step-ldt
  "Shift a LocalDateTime by k units (k may be 0); handles every Org repeater unit."
  [ldt k unit]
  (case unit
    :hour  (.plusHours ldt k)
    :day   (.plusDays ldt k)
    :week  (.plusWeeks ldt k)
    :month (.plusMonths ldt k)
    :year  (.plusYears ldt k)
    nil))

(def ^:private chrono-units
  {:hour  java.time.temporal.ChronoUnit/HOURS
   :day   java.time.temporal.ChronoUnit/DAYS
   :week  java.time.temporal.ChronoUnit/WEEKS
   :month java.time.temporal.ChronoUnit/MONTHS
   :year  java.time.temporal.ChronoUnit/YEARS})

(defn- repeating?
  "True for a repeater that actually repeats: a positive count and a known
   unit (organ yields {:n 0 ...} for a +0d cookie, which would divide by 0)."
  [{:keys [n unit]}]
  (boolean (and n (pos? n) (chrono-units unit))))

(defn- ics-window
  "Optional [win-s win-e] LocalDateTime bounds from --reference-date/--weeks,
   nil when neither is set. --reference-date alone leaves the window
   open-ended; --weeks alone starts it today (in tz)."
  [opts tz]
  (when (or (:weeks opts) (:reference-date opts))
    (let [ref (or (:reference-date opts) (java.time.LocalDate/now tz))]
      [(.atStartOfDay ref)
       (when-let [w (:weeks opts)]
         (.atStartOfDay (.plusWeeks ref w)))])))

(defn- iso->ldt
  "LocalDateTime of an ISO timestamp (2025-01-15T10:30, or 2025-01-15 which
   counts from midnight)."
  [s]
  (if (str/includes? s "T")
    (java.time.LocalDateTime/parse s)
    (.atStartOfDay (java.time.LocalDate/parse s))))

(defn- opts->zone
  "Zone anchoring the Org timestamps: --time-zone, else the system default."
  [opts]
  (or (:time-zone opts) (java.time.ZoneId/systemDefault)))

(defn- item-in-window?
  "True when the timestamp entry starts inside [win-s win-e) — for
   repeaters, when any occurrence does. A nil win-e means no upper bound."
  [{:keys [start repeater]} win-s win-e]
  (let [start            (iso->ldt start)
        {:keys [n unit]} repeater
        before-end?      (fn [ldt] (or (nil? win-e) (.isBefore ldt win-e)))]
    (if (.isBefore start win-s)
      ;; starts before the window: only a repeater can reach it
      (boolean
       (when (repeating? repeater)
         (let [gap (.between (chrono-units unit) start win-s)
               k0  (* n (quot gap n))]
           (->> (iterate #(+ % n) k0)
                (map #(step-ldt start % unit))
                (drop-while #(.isBefore % win-s))
                first
                before-end?))))
      (before-end? start))))

(defn- collect-ics-items
  "SCHEDULED -> VEVENT and DEADLINE+TODO -> VTODO entries from organ's
   active-timestamps (inline timestamps are not exported), each tagged with
   an :ics-type. DONE entries are pruned upstream unless --keep-done."
  [ast]
  (->> (organ/active-timestamps ast)
       (keep (fn [{:keys [origin todo] :as entry}]
               (case origin
                 :scheduled (assoc entry :ics-type :vevent)
                 :deadline  (when todo (assoc entry :ics-type :vtodo))
                 nil)))))

(defn- uid-for-item
  "Deterministic UID from the item's own properties, not its position in
   the document — reordering entries must not churn the UIDs calendar
   clients rely on. dup counts identical (title, timestamp, type) items."
  [item dup]
  (let [input (str/join "\n" [(:title item) (:value item)
                              (name (:ics-type item)) dup])
        bytes (.digest (java.security.MessageDigest/getInstance "SHA-256")
                       (.getBytes input "UTF-8"))
        hex   (str/join (map #(format "%02x" %) bytes))]
    (str (subs hex 0 16) "@hop")))

(defn- assign-uids
  "Assoc a :uid on each item, counting duplicates in document order."
  [items]
  (:acc (reduce (fn [{:keys [seen acc]} item]
                  (let [k (mapv item [:title :value :ics-type])
                        n (get seen k 0)]
                    {:seen (assoc seen k (inc n))
                     :acc  (conj acc (assoc item :uid (uid-for-item item n)))}))
                {:seen {} :acc []}
                items)))

(def ^:private ics-rrule-freq
  {:hour "HOURLY" :day "DAILY" :week "WEEKLY" :month "MONTHLY" :year "YEARLY"})

(defn- org-repeater-to-rrule
  "Convert a parsed Org repeater ({:n :unit :kind}, as produced by
   organ/active-timestamps) to an ICS RRULE string. The shift kind (+/++/.+)
   has no RRULE equivalent and is ignored."
  [{:keys [n unit] :as repeater}]
  (when (repeating? repeater)
    (str "RRULE:FREQ=" (ics-rrule-freq unit) ";INTERVAL=" n)))

(defn- ics-next-day
  "Given an ICS VALUE=DATE string like '20250115', return the next day '20250116'.
   RFC 5545 requires DTEND to be exclusive for VALUE=DATE events."
  [date-str]
  (-> (java.time.LocalDate/parse date-str
                                 (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd"))
      (.plusDays 1)
      (.format (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd"))))

(def ^:private ics-local-dtf
  "ICS floating datetime formatter (20250115T103000)."
  (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd'T'HHmmss"))

(defn- ics-utc
  "UTC ICS datetime (20250115T093000Z) of a floating ICS datetime
   (20250115T103000) anchored in tz."
  [s tz]
  (.format utc-dtf (.toInstant (.atZone (java.time.LocalDateTime/parse s ics-local-dtf) tz))))

(defn- vcalendar
  "Wrap rendered components in a VCALENDAR (CRLF line endings), with extra
   header lines after PRODID."
  [header-lines components]
  (str (str/join "\r\n"
                 (concat ["BEGIN:VCALENDAR" "VERSION:2.0" "PRODID:-//hop//EN"]
                         header-lines
                         components
                         ["END:VCALENDAR"]))
       "\r\n"))

(defn- ics-component
  "Render one item as a VEVENT (scheduled) or VTODO (deadline) component."
  [{:keys [ics-type title todo value repeater uid]} done-keywords opts tz now]
  (let [{:keys [dtstart dtend]} (iso-to-ics-datetime value)
        is-date (not (str/includes? dtstart "T"))
        tag     (if (= ics-type :vevent) "VEVENT" "VTODO")
        summary (if (and todo (= ics-type :vtodo)) (str (name todo) " " title) title)
        dates   (case ics-type
                  :vevent (if is-date
                            [(str "DTSTART;VALUE=DATE:" dtstart)
                             (str "DTEND;VALUE=DATE:" (ics-next-day dtstart))]
                            (let [dtend (or dtend
                                            (when-let [d (:default-duration opts)]
                                              (-> (java.time.LocalDateTime/parse dtstart ics-local-dtf)
                                                  (.plusMinutes d)
                                                  (.format ics-local-dtf))))]
                              (cond-> [(str "DTSTART:" (ics-utc dtstart tz))]
                                dtend (conj (str "DTEND:" (ics-utc dtend tz))))))
                  :vtodo  [(str "DUE" (when is-date ";VALUE=DATE") ":"
                                (if is-date dtstart (ics-utc dtstart tz)))])
        status  (when (= ics-type :vtodo)
                  [(str "STATUS:" (if (contains? done-keywords (name todo))
                                    "COMPLETED" "NEEDS-ACTION"))])]
    (str/join "\r\n"
              (concat [(str "BEGIN:" tag)
                       (ics-fold-line (str "UID:" uid))
                       (str "DTSTAMP:" now)]
                      dates
                      (when-let [rrule (org-repeater-to-rrule repeater)] [rrule])
                      [(ics-fold-line (str "SUMMARY:" (ics-escape summary)))]
                      status
                      [(str "END:" tag)]))))

(defn- render-ast-as-ics
  "Export scheduled items as VEVENT and deadline+TODO items as VTODO.
   With --keep-done, DONE deadlines yield VTODOs with STATUS:COMPLETED.
   All-day entries are excluded unless --all-day. --reference-date/--weeks
   bound the exported items (a repeater keeps its RRULE when an occurrence
   reaches the window); --default-duration fills in missing end times.
   Datetimes are anchored in --time-zone (system timezone by default) and
   emitted as UTC."
  [ast done-keywords opts]
  (let [tz            (opts->zone opts)
        [win-s win-e] (ics-window opts tz)
        now           (.format utc-dtf (java.time.Instant/now))
        items         (cond->> (collect-ics-items ast)
                        (not (:all-day opts)) (remove :all-day)
                        win-s (filter #(item-in-window? % win-s win-e))
                        true  assign-uids)]
    (vcalendar [] (map #(ics-component % done-keywords opts tz now) items))))

;; ICS anonymised (free/busy) export -- the `ics-anon` format
;; ---------------------------------------------------------------------------
;; Re-emits SCHEDULED timestamps as anonymised "Busy" blocks — the VEVENT
;; side of the plain `ics` format. DEADLINE timestamps are due dates (its
;; VTODO side) and never count as busy.
;; With --weeks, repeaters are expanded over the bounded window and strictly
;; overlapping intervals are merged (touching boundaries stay distinct);
;; without it every timestamp yields one block and repeaters keep their
;; RRULE, like the plain `ics` format. All-day timestamps are skipped
;; unless --all-day, which only admits those without a TODO keyword
;; (e.g. a "Vacation" heading blocks its day; a task scheduled on a date
;; does not). --time-zone only anchors the org timestamps; output
;; datetimes are normalised to UTC (Z suffix) so no VTIMEZONE is needed.
;; The plain `ics` format anchors and emits its datetimes the same way.

(defn- ldt->zdt [ldt tz] (java.time.ZonedDateTime/of ldt tz))
(defn- zmax [a b] (if (.isAfter a b) a b))

(defn- entry->base
  "Base [start end] LocalDateTimes of a timestamp entry (--default-duration
   fills a missing end; without it the interval is zero-length)."
  [{:keys [start end all-day]} event-duration]
  (if all-day
    (let [s (iso->ldt start)]
      [s (.plusDays s 1)])
    (let [s    (iso->ldt start)
          fill #(if event-duration (.plusMinutes s event-duration) s)
          e    (if end
                 (let [e0 (java.time.LocalDateTime/parse end)]
                   (cond (.isAfter e0 s)  e0
                         ;; end before start: overnight range (23:00-01:00)
                         (.isBefore e0 s) (.plusDays e0 1)
                         :else            (fill)))
                 (fill))]
      [s e])))

(defn- first-repeat-offset
  "Offset (in repeater units, multiple of n) of the first occurrence that may
   reach a window starting at win-s-ldt; 0 when the series starts inside or
   after the window. Computed analytically so decades-old repeaters don't
   exhaust the occurrence guard before reaching the window."
  [start-ldt end-ldt win-s-ldt n unit]
  (let [cu       (get chrono-units unit)
        duration (.between cu start-ldt end-ldt)
        gap      (.between cu start-ldt win-s-ldt)]
    (* n (max 0 (dec (quot (- gap duration) n))))))

(defn- entry->intervals
  "Zoned [start end] intervals of an entry, expanded over the window by its repeater."
  [entry event-duration tz win-s win-e]
  (let [[start-ldt end-ldt] (entry->base entry event-duration)
        {:keys [n unit]}    (:repeater entry)
        in-window?          (fn [zs ze] (and (.isBefore zs win-e) (.isAfter ze win-s)))]
    (if (repeating? (:repeater entry))
      (let [k0 (first-repeat-offset start-ldt end-ldt (.toLocalDateTime win-s) n unit)]
        (->> (iterate #(+ % n) k0)
             (take 5000) ; safety guard
             (map (fn [k] [(ldt->zdt (step-ldt start-ldt k unit) tz)
                           (ldt->zdt (step-ldt end-ldt k unit) tz)]))
             (take-while (fn [[zs _]] (.isBefore zs win-e)))
             (filter (fn [[zs ze]] (in-window? zs ze)))))
      (let [zs (ldt->zdt start-ldt tz)
            ze (ldt->zdt end-ldt tz)]
        (when (in-window? zs ze) [[zs ze]])))))

(defn- busy-entries
  "SCHEDULED timestamp entries eligible for busy output — the VEVENT side
   of the plain ics format. DEADLINE entries are due dates (the VTODO
   side): they never reserve time, so they never count as busy. DONE
   entries are pruned upstream unless --keep-done. All-day entries are
   skipped by default; with --all-day, only those whose heading bears no
   TODO keyword count as busy — a keyword marks a task parked on a date,
   not a day-blocking event."
  [ast all-day?]
  (->> (organ/active-timestamps ast)
       (filter #(= :scheduled (:origin %)))
       (filter #(or (not (:all-day %))
                    (and all-day? (not (:todo %)))))))

(defn- sort-intervals [ivs]
  (sort-by #(.toInstant (first %)) ivs))

(defn- merge-intervals-zoned
  "Merge zoned intervals that strictly overlap (sorted input). Touching
   boundaries (e.g. back-to-back all-day events at midnight) stay distinct so
   each occurrence renders as its own VEVENT."
  [ivs]
  (reduce (fn [acc [s e]]
            (if (and (seq acc) (.isBefore s (second (peek acc))))
              (conj (pop acc) [(first (peek acc)) (zmax (second (peek acc)) e)])
              (conj acc [s e])))
          [] ivs))

(defn- render-busy-ics
  "Emit a VCALENDAR of zoned [start end rrule?] intervals as anonymised Busy
   VEVENTs. Datetimes are normalised to UTC (Z suffix) so no VTIMEZONE is
   needed; zero-length intervals get no DTEND."
  [intervals]
  (let [now     (.format utc-dtf (java.time.Instant/now))
        fmt-utc (fn [zdt] (.format utc-dtf (.toInstant zdt)))
        event   (fn [idx [s e rrule]]
                  (str/join "\r\n"
                            (concat
                             ["BEGIN:VEVENT"
                              (ics-fold-line (str "UID:busy-" idx "-" (fmt-utc s) "@hop"))
                              (str "DTSTAMP:" now)
                              (str "DTSTART:" (fmt-utc s))]
                             (when-not (.isEqual s e)
                               [(str "DTEND:" (fmt-utc e))])
                             (when rrule [rrule])
                             [(ics-fold-line (str "SUMMARY:" (ics-escape "Busy")))
                              "STATUS:CONFIRMED"
                              "TRANSP:OPAQUE"
                              "END:VEVENT"])))]
    (vcalendar ["CALSCALE:GREGORIAN"] (map-indexed event intervals))))

(defn- render-ast-as-busy-ics
  "Render the AST as anonymised \"Busy\" blocks. With --weeks, repeaters are
   expanded over the bounded window and overlapping blocks are merged;
   without it every timestamp yields one block and repeaters keep their
   RRULE, like the plain ics format. --reference-date filters out items
   starting before it."
  [ast opts]
  (let [tz       (opts->zone opts)
        duration (:default-duration opts)
        ref      (or (:reference-date opts) (java.time.LocalDate/now tz))
        entries  (busy-entries ast (:all-day opts))]
    (if-let [w (:weeks opts)]
      (let [win-s  (ldt->zdt (.atStartOfDay ref) tz)
            win-e  (ldt->zdt (.atStartOfDay (.plusWeeks ref w)) tz)
            events (mapcat #(entry->intervals % duration tz win-s win-e) entries)]
        (render-busy-ics (-> events sort-intervals merge-intervals-zoned)))
      (let [win-s  (when (:reference-date opts) (.atStartOfDay ref))
            blocks (->> entries
                        (filter #(or (nil? win-s) (item-in-window? % win-s nil)))
                        (map (fn [entry]
                               (let [[s e] (entry->base entry duration)]
                                 [(ldt->zdt s tz) (ldt->zdt e tz)
                                  (org-repeater-to-rrule (:repeater entry))])))
                        sort-intervals)]
        (render-busy-ics blocks)))))

(defn- prepare-ast [ast]
  (let [options (parse-options-string (meta-string ast :options))]
    (-> ast
        (assoc :parsed-options options)
        number-sections
        assign-section-ids)))
(defn render-ast-as-markdown [ast] (render-node (prepare-ast ast) :md))
(defn render-ast-as-html [ast] (render-node (prepare-ast ast) :html))
(defn render-ast-as-org [ast] (render-node (prepare-ast ast) :org))

;; Output Formatting
(defn format-ast-as-json [ast] (json/generate-string ast {:pretty true}))
(defn format-ast-as-yaml [ast]
  (yaml/generate-string ast {:dumper-options {:flow-style :block}}))

;; Statistics
(defn- count-words
  "Count words in a string or vector of inline nodes."
  [s]
  (let [^String text (cond (string? s)     s
                           (sequential? s) (inline-text s)
                           :else           nil)]
    (if (non-blank? text)
      (let [len (.length text)]
        (loop [i 0 in-word false n 0]
          (if (>= i len)
            (if in-word (inc n) n)
            (let [ws (Character/isWhitespace (.charAt text i))]
              (recur (inc i) (not ws) (if (and in-word ws) (inc n) n))))))
      0)))

(defn- count-images
  "Count image links in a vector of inline nodes (0 for strings)."
  [nodes]
  (if (sequential? nodes)
    (reduce (fn [n node]
              (+ n (case (:type node)
                     :link (if (image-url? (or (:target node) (:url node))) 1 0)
                     (:bold :italic :underline :strike :footnote-inline)
                     (count-images (:children node))
                     0)))
            0 nodes)
    0))

(defn- content-stats
  "Extract word and image counts from a node's :content field."
  [node]
  {:words  (count-words (:content node))
   :images (count-images (:content node))})

(defn- inc-stat
  "Increment counter k in m (absent counts as 0)."
  [m k]
  (update m k (fnil inc 0)))

(defn- add-stat
  "Add n to counter k in m (absent counts as 0)."
  [m k n]
  (update m k (fnil + 0) n))

(defn- collect-stats
  "Recursively collect statistics from AST nodes."
  [node]
  (let [node-type   (:type node)
        children    (:children node [])
        items       (:items node [])
        child-stats (reduce (fn [acc child]
                              (merge-with + acc (collect-stats child)))
                            {}
                            (concat children items))
        base-stats  (case node-type
                      :section         (let [planning (:planning node)]
                                         (cond-> (-> child-stats
                                                     (inc-stat :sections)
                                                     (add-stat :words (count-words (:title node))))
                                           (:scheduled planning) (inc-stat :scheduled)
                                           (:deadline planning)  (inc-stat :deadline)
                                           (:closed planning)    (inc-stat :closed)))
                      (:paragraph :quote-block :footnote-def)
                      (let [{:keys [words images]} (content-stats node)
                            stat-key               (case node-type
                                                     :paragraph    :paragraphs
                                                     :quote-block  :quote-blocks
                                                     :footnote-def :footnotes)]
                        (-> child-stats
                            (inc-stat stat-key)
                            (add-stat :words words)
                            (add-stat :images images)))
                      :table           (inc-stat child-stats :tables)
                      :list            (inc-stat child-stats :lists)
                      :list-item       (-> child-stats
                                           (inc-stat :list-items)
                                           (add-stat :words (+ (count-words (:content node))
                                                               (count-words (:term node))
                                                               (count-words (:definition node))))
                                           (add-stat :images (+ (count-images (:content node))
                                                                (count-images (:definition node)))))
                      :src-block       (inc-stat child-stats :src-blocks)
                      :block           (inc-stat child-stats :blocks)
                      :comment         (inc-stat child-stats :comments)
                      :fixed-width     (inc-stat child-stats :fixed-width)
                      :html-line       (inc-stat child-stats :html-lines)
                      :latex-line      (inc-stat child-stats :latex-lines)
                      :property-drawer (inc-stat child-stats :property-drawers)
                      :document        child-stats
                      child-stats)]
    base-stats))

(defn compute-stats
  "Compute statistics for an AST and return a sorted map."
  [ast]
  (let [raw-stats     (collect-stats ast)
        ordered-keys  [:sections :paragraphs :words :images :lists :list-items
                       :tables :src-blocks :quote-blocks :blocks
                       :footnotes :comments :fixed-width :html-lines
                       :latex-lines :property-drawers
                       :scheduled :deadline :closed]
        index-map     (zipmap ordered-keys (range))
        present-stats (into {} (filter (fn [[_ v]] (and v (pos? v))) raw-stats))]
    (into (sorted-map-by (fn [a b]
                           (let [c (compare (get index-map a 999)
                                            (get index-map b 999))]
                             (if (zero? c) (compare a b) c))))
          present-stats)))

(def ^:private stats-label-map
  {:sections         "Sections"
   :paragraphs       "Paragraphs"
   :words            "Words"
   :images           "Images"
   :lists            "Lists"
   :list-items       "List items"
   :tables           "Tables"
   :src-blocks       "Source blocks"
   :quote-blocks     "Quote blocks"
   :blocks           "Other blocks"
   :footnotes        "Footnotes"
   :comments         "Comments"
   :fixed-width      "Fixed-width blocks"
   :html-lines       "HTML lines"
   :latex-lines      "LaTeX lines"
   :property-drawers "Property drawers"
   :scheduled        "Scheduled"
   :deadline         "Deadline"
   :closed           "Closed"})

(defn format-stats
  "Format stats as aligned 'Label: value' lines."
  [stats]
  (let [max-label-len (->> (keys stats)
                           (map #(count (get stats-label-map % (name %))))
                           (reduce max 0))]
    (->> stats
         (map (fn [[k v]]
                (let [label   (get stats-label-map k (name k))
                      padding (repeat-str (- max-label-len (count label)) " ")]
                  (str label ": " padding v))))
         (str/join "\n"))))

;; CLI Options

(defn- resolve-css-theme
  "Resolve a CSS theme value to a map with :link (URL) or :inline (CSS content).
  Resolution order:
  1. https:// URL → {:link url}
  2. file:/// URL → {:inline content}; must point to an existing .css file
  3. .css filename matching a local file → {:inline content};
     if the file doesn't exist, throw (no CDN fall-back)
  4. bare name (no spaces) → pico-themes CDN {:link url}"
  [v]
  (cond
    (re-matches #"https?://.*" v)
    {:link v}

    (str/starts-with? v "file:///")
    (let [path (subs v 7)]
      (if-not (str/ends-with? path ".css")
        (throw (ex-info (str "CSS theme must point to a .css file: " v) {:theme v}))
        (if (.exists (java.io.File. path))
          {:inline (slurp path)}
          (throw (ex-info (str "CSS theme file not found: " path) {:theme v})))))

    (str/ends-with? v ".css")
    (if (.exists (java.io.File. v))
      {:inline (slurp v)}
      (throw (ex-info (str "CSS theme file not found: " v) {:theme v})))

    (re-matches #"\S+" v)
    {:link (str pico-themes-cdn v ".css")}

    :else
    (throw (ex-info (str "CSS theme not found: " v
                         " (expected a URL, a local .css file, or a pico-theme name)")
                    {:theme v}))))

(def ^:private cli-options
  [["-b" "--base-url URL" "Base URL prepended to relative links (include trailing slash)"]
   ["-c" "--css-theme THEME" "CSS theme: a URL, a local .css file, or a pico-theme name (e.g. org, doric)"]
   ["-f" "--format FORMAT" "Output format: json, edn, yaml, md, html, org, ics, or ics-anon"
    :default "json" :validate [#{"json" "edn" "yaml" "md" "html" "org" "ics" "ics-anon"} "Must be: json, edn, yaml, md, html, org, ics, ics-anon"]]
   ["-h" "--help" "Show help"]
   ["-v" "--version" "Show the hop version"]
   ["-i" "--id REGEX" "Filter: section's own ID or CUSTOM_ID matches" :parse-fn re-pattern]
   ["-I" "--under-id REGEX" "Filter: any ancestor's ID or CUSTOM_ID matches (keeps the whole subtree under that ancestor)" :parse-fn re-pattern]
   ["-l" "--level-limit LEVEL" "Filter: level <= LEVEL"
    :parse-fn #(Integer/parseInt %) :validate [pos? "Must be positive"]]
   ["-L" "--level-limit-inclusive LEVEL" "Filter: level <= LEVEL and render deeper headings as bold"
    :parse-fn #(Integer/parseInt %) :validate [pos? "Must be positive"]]
   ["-n" "--no-unwrap" "Preserve original line breaks"]
   ["-r" "--render FORMAT" "Content rendering format in AST output: md, html, or org"
    :default "md" :validate [#{"md" "html" "org"} "Must be: md, html, org"]]
   ["-s" "--stats" "Compute and display document statistics"]
   ["-t" "--title REGEX" "Filter: section's own title matches" :parse-fn re-pattern]
   ["-T" "--under-title REGEX" "Filter: any ancestor's title matches (keeps the whole subtree under that ancestor)" :parse-fn re-pattern]
   ;; Cross-format DONE handling
   ["-D" "--done-keywords KW,KW" "Comma-separated extra DONE keywords (e.g. CANX,CANCELED); merged with DONE and any #+TODO: directives in the file"
    :parse-fn #(into #{} (remove str/blank? (map str/trim (str/split % #","))))]
   ["-k" "--keep-done" "Keep subtrees whose heading is in a DONE state (by default they're stripped from every output)"]
   ;; Tuning for the ics and ics-anon formats
   ["-a" "--all-day" "also emit all-day SCHEDULED/DEADLINE (excluded by default; ics-anon only takes those without a TODO keyword)"]
   ["-d" "--default-duration N" "minutes for events without an end time (only added when set)"
    :parse-fn #(Integer/parseInt %) :validate [pos? "Must be positive"]]
   ["-w" "--weeks N" "weeks ahead to scan (unbounded unless set; a bounded window lets ics-anon expand and merge repeaters)"
    :parse-fn #(Integer/parseInt %) :validate [pos? "Must be positive"]]
   ["-z" "--time-zone ZONE" "timezone anchoring the org timestamps (default: system timezone); output datetimes are UTC"
    :parse-fn #(java.time.ZoneId/of %)]
   ["-R" "--reference-date DATE" "window start date, YYYY-MM-DD (today when only -w is set; useful for reproducible output)"
    :parse-fn #(java.time.LocalDate/parse %)]])

(defn- bbin-version
  "Version string read from bbin's install metadata (`bbin ls --edn`), or nil
  when bbin is absent or hop was not installed through it."
  []
  (try
    (let [{:keys [out exit]} (process/shell {:out :string :err :string :continue true}
                                            "bbin" "ls" "--edn")]
      (when (zero? exit)
        (some (fn [e] (when (= 'io.github.bzg/hop (:lib e))
                        (get-in e [:coords :git/tag])))
              (vals (edn/read-string out)))))
    (catch Exception _ nil)))

(defn- usage [summary]
  (str/join \newline
            ["Hush Org Parser - Render and export Org files"
             "" "Usage: hop [options] <org-file>"
             "" "Options:" summary]))

(defn- exit-error
  "Print error message to stderr and exit with code 1."
  ([msg] (exit-error msg nil))
  ([msg summary]
   (binding [*out* *err*]
     (println "Error:" msg)
     (when summary (println (usage summary))))
   (System/exit 1)))

(defn- exit-ok
  "Print message to stdout and exit with code 0."
  [msg]
  (println msg) (System/exit 0))

(defn- filter-opts
  "organ/filter-ast options from the parsed CLI options."
  [options]
  {:level-limit           (:level-limit options)
   :level-limit-inclusive (:level-limit-inclusive options)
   :title-pattern         (:title options)
   :id-pattern            (:id options)
   :section-title-pattern (:under-title options)
   :section-id-pattern    (:under-id options)})

(defn- report-parse-errors!
  "Print organ's parse errors, if any, as warnings on stderr."
  [ast]
  (when-let [errs (:parse-errors ast)]
    (binding [*out* *err*]
      (doseq [{:keys [line message]} errs]
        (println (str "Warning (line " line "): " message))))))

(defn- render-output
  "Render the parsed, filtered AST in the requested output format.
   Org output keeps the raw AST; every other format works on the cleaned
   one (trimmed content, blank nodes removed). ICS strings are CRLF
   terminated and must be printed as is."
  [ast done-keywords options]
  (let [fmt     (:format options)
        cleaned (if (= fmt "org") ast (organ/clean-node ast))]
    (report-parse-errors! cleaned)
    (if (:stats options)
      (format-stats (compute-stats ast))
      (case fmt
        "md"       (render-ast-as-markdown cleaned)
        "html"     (render-ast-as-html cleaned)
        "org"      (render-ast-as-org cleaned)
        "ics"      (render-ast-as-ics ast done-keywords options)
        "ics-anon" (render-ast-as-busy-ics ast options)
        (let [rendered (render-content-in-node cleaned (keyword (:render options)))]
          (case fmt
            "json" (format-ast-as-json rendered)
            "edn"  (organ/format-ast-as-edn rendered)
            "yaml" (format-ast-as-yaml rendered)))))))

(defn- run
  "Parse org-content, apply the CLI filters and print the output."
  [org-content options]
  (binding [*base-url*  (when-let [u (:base-url options)]
                          (cond-> u (not (str/ends-with? u "/")) (str "/")))
            *css-theme* (some-> (:css-theme options) resolve-css-theme)]
    (let [ast           (organ/parse-org
                         org-content
                         {:unwrap?       (not (:no-unwrap options))
                          :todo-keywords {:done (vec (:done-keywords options))}})
          done-keywords (set (:done (:todo-keywords ast)))
          filtered      (cond-> (organ/filter-ast ast (filter-opts options))
                          (not (:keep-done options)) (prune-done-sections done-keywords))
          output        (render-output filtered done-keywords options)]
      (if (str/starts-with? (:format options) "ics")
        (print output)
        (println output))
      (flush))))

(defn -main [& args]
  (let [{:keys [options arguments errors summary]} (cli/parse-opts args cli-options)]
    (cond
      (:version options)
      (exit-ok (str "hop " (or (bbin-version) "(version unknown)")))

      (:help options)
      (exit-ok (usage summary))

      errors
      (exit-error (str/join "; " errors) summary)

      (not= (count arguments) 1)
      (exit-error "Expected one argument (org-file or '-')" summary)

      :else
      (let [file-path   (first arguments)
            org-content (if (= file-path "-")
                          (slurp *in*)
                          (try (slurp file-path)
                               (catch java.io.FileNotFoundException _
                                 (exit-error (str "File not found - " file-path)))))]
        (try
          (run org-content options)
          (System/exit 0)
          (catch clojure.lang.ExceptionInfo e
            ;; Expected user-facing errors (e.g. an unusable --css-theme)
            (exit-error (.getMessage e)))
          (catch Exception e
            (binding [*out* *err*]
              (println "Error:" (.getMessage e))
              (.printStackTrace e *err*))
            (System/exit 1)))))))

(when (= *file* (System/getProperty "babashka.file"))
  (apply -main *command-line-args*))
