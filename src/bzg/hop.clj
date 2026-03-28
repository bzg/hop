#!/usr/bin/env bb

;; Copyright (c) Bastien Guerry
;; SPDX-License-Identifier: EPL-2.0
;; License-Filename: LICENSES/EPL-2.0.txt

(ns bzg.hop
  "Hush Org Parser - Render and export Org ASTs produced by organ."
  (:require [clojure.string :as str]
            [cheshire.core :as json]
            [clojure.tools.cli :as cli]
            [clj-yaml.core :as yaml]
            [bzg.organ :as organ]))

;; Re-bind organ utilities for local use
(def ^:private remove-empty-vals organ/remove-empty-vals)
(def ^:private non-blank? organ/non-blank?)
(def ^:private escape-block-content organ/escape-block-content)
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
;; Organ 0.2.0 provides parsed inline nodes. These functions render them
;; to HTML, Markdown, or Org format strings.

(defn- escape-html [text]
  (-> (str text)
      (str/replace "&" "&amp;")
      (str/replace "<" "&lt;")
      (str/replace ">" "&gt;")
      (str/replace "\"" "&quot;")))

(defn- upper-name [k] (str/upper-case (name k)))

(defn- repeat-str
  "Repeat string s exactly n times. Returns empty string when n <= 0."
  [n s]
  (str/join (repeat n s)))

(defn- prefix-lines
  "Prefix each line of content with the given string."
  [prefix content]
  (->> (str/split-lines content)
       (map #(str prefix %))
       (str/join "\n")))

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

(defn- remote-url?
  "Check if URL is a remote HTTP/HTTPS URL."
  [url]
  (when url
    (boolean (re-find #"^https?://" url))))

(defn- expand-home
  "Expand ~ to user home directory in file paths."
  [path]
  (if (and path (str/starts-with? path "~/"))
    (str (System/getProperty "user.home") (subs path 1))
    path))

(defn- file-to-base64
  "Read a file and return its base64-encoded content, or nil if file doesn't exist."
  [filepath]
  (try
    (let [file (java.io.File. (expand-home filepath))]
      (when (.exists file)
        (-> file .toPath java.nio.file.Files/readAllBytes
            (->> (.encodeToString (java.util.Base64/getEncoder))))))
    (catch Exception _ nil)))

(defn- image-to-data-uri
  "Convert a local image file to a data URI, or nil if not possible."
  [filepath]
  (when-let [mime (some-> (get-file-extension filepath) image-mime-types)]
    (when-let [b64 (file-to-base64 filepath)]
      (str "data:" mime ";base64," b64))))

(defn- build-img-attrs
  "Build HTML attribute string from a map of attributes.
   Handles :alt, :title, :width, :height, :class, :id, :style, etc."
  [attrs]
  (when (seq attrs)
    (->> attrs
         (keep (fn [[k v]]
                 (when (non-blank? (str v))
                   (str " " (name k) "=\"" (escape-html (str v)) "\""))))
         (str/join))))

(defn- render-image-html
  "Render an image tag with optional affiliated attributes.
   src: the image URL or data URI
   default-alt: fallback alt text if not specified in attrs
   affiliated: the affiliated keywords map (may contain :caption, :attr {:html {...}})"
  [src default-alt affiliated]
  (let [html-attrs (get-in affiliated [:attr :html] {})
        ;; Use alt from #+attr_html if provided, otherwise default
        alt (or (:alt html-attrs) default-alt)
        ;; Build the img tag with all HTML attributes
        img-attrs (merge {:src src :alt alt}
                         (dissoc html-attrs :alt)) ; alt already handled
        img-tag (str "<img" (build-img-attrs img-attrs) ">")
        ;; Wrap in figure with figcaption if caption is present
        caption (:caption affiliated)]
    (if caption
      (str "<figure>" img-tag "<figcaption>"
           (escape-html caption) "</figcaption></figure>")
      img-tag)))

(defn- heading-to-slug
  "Convert a heading title to a URL-safe slug for anchor links.
   Accepts both a string or a vector of inline nodes."
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
  (if (and *base-url*
           (not (str/starts-with? href "/"))
           (not (str/starts-with? href "#"))
           (not (re-find #"^[a-zA-Z][a-zA-Z0-9+.-]*:" href)))
    (str *base-url* href)
    href))

(defn- resolve-link-href
  "Resolve a link node to a URL string given its fields."
  [{:keys [link-type target url]}]
  (prepend-base-url
   (case link-type
     :file          (str/replace target #"\.org$" ".html")
     (:id
      :custom-id)  (str "#" target)
     :heading       (str "#" (heading-to-slug target))
     :mailto        (str "mailto:" target)
     url)))

;; Inline Node Rendering — forward declarations
(declare render-inline render-link-node)

(defn- render-inline-node
  "Render a single inline node to a string in the given format."
  [node fmt]
  (case (:type node)
    :text      (case fmt :html (escape-html (:value node)) (:value node))
    :bold      (case fmt
                 :html (str "<strong>" (render-inline (:children node) fmt) "</strong>")
                 :md   (str "**" (render-inline (:children node) fmt) "**")
                 (str "*" (render-inline (:children node) fmt) "*"))
    :italic    (case fmt
                 :html (str "<em>" (render-inline (:children node) fmt) "</em>")
                 :md   (str "*" (render-inline (:children node) fmt) "*")
                 (str "/" (render-inline (:children node) fmt) "/"))
    :underline (case fmt
                 :html (str "<u>" (render-inline (:children node) fmt) "</u>")
                 :md   (str "_" (render-inline (:children node) fmt) "_")
                 (str "_" (render-inline (:children node) fmt) "_"))
    :strike    (case fmt
                 :html (str "<del>" (render-inline (:children node) fmt) "</del>")
                 :md   (str "~~" (render-inline (:children node) fmt) "~~")
                 (str "+" (render-inline (:children node) fmt) "+"))
    :code      (case fmt
                 :html (str "<code>" (escape-html (:value node)) "</code>")
                 :md   (str "`" (:value node) "`")
                 (str "~" (:value node) "~"))
    :verbatim  (case fmt
                 :html (str "<code>" (escape-html (:value node)) "</code>")
                 :md   (str "`" (:value node) "`")
                 (str "=" (:value node) "="))
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
        (str "[fn:" (or label "") ":" (or (inline-text children) "") "]")))
    :timestamp (:raw node)
    ;; fallback
    (or (:value node) "")))

(defn- render-inline
  "Render a vector of inline nodes to a string in the given format.
   If given a string (fallback), escapes for HTML or returns as-is."
  [nodes fmt]
  (cond
    (nil? nodes) ""
    (string? nodes) (case fmt :html (escape-html nodes) nodes)
    (empty? nodes) ""
    :else (str/join (map #(render-inline-node % fmt) nodes))))

(defn- extract-single-image-node
  "If an inline node vector contains exactly one link node that is an image,
   return that node. Otherwise nil."
  [nodes]
  (when (and (sequential? nodes)
             (= 1 (count nodes))
             (= :link (:type (first nodes))))
    (let [node (first nodes)
          actual-url (if (#{:http :https} (:link-type node))
                       (:url node) (:target node))]
      (when (image-url? actual-url) node))))

(defn- render-link-node
  "Render a :link inline node. Optionally accepts affiliated keywords for
   enhanced image rendering (e.g., paragraphs with #+CAPTION)."
  [node fmt affiliated]
  (let [link-type   (:link-type node)
        target      (:target node)
        url         (:url node)
        children    (:children node)
        has-desc    (seq children)
        desc-text   (when has-desc (inline-text children))
        actual-url  (if (#{:http :https} link-type) url (or target url))
        is-local-file (= link-type :file)
        is-remote   (#{:http :https} link-type)
        url-is-image  (image-url? actual-url)
        desc-is-image (and desc-text (image-url? desc-text))]
    (cond
      ;; Local file images: convert to base64 data URI
      (and is-local-file url-is-image)
      (if-let [data-uri (image-to-data-uri (or target url))]
        (cond
          (= fmt :html)
          (if affiliated
            (render-image-html data-uri (or desc-text target) affiliated)
            (if desc-is-image
              (str "<a href=\"" data-uri "\"><img src=\""
                   (escape-html desc-text) "\" alt=\"" (escape-html desc-text) "\"></a>")
              (str "<img src=\"" data-uri "\" alt=\""
                   (escape-html (or desc-text target)) "\">")))
          (= fmt :md)
          (str "![" (or desc-text target) "](" data-uri ")")
          :else "")
        "")

      ;; HTML format
      (= fmt :html)
      (cond
        (and is-remote url-is-image (not has-desc))
        (if affiliated
          (render-image-html url url affiliated)
          (str "<img src=\"" (escape-html url) "\" alt=\""
               (escape-html url) "\">"))

        (and is-remote url-is-image has-desc (not desc-is-image))
        (if affiliated
          (render-image-html url desc-text affiliated)
          (str "<img src=\"" (escape-html url) "\" alt=\""
               (escape-html desc-text) "\">"))

        (and is-remote desc-is-image)
        (str "<a href=\"" (escape-html url) "\"><img src=\""
             (escape-html desc-text) "\" alt=\"" (escape-html desc-text) "\"></a>")

        :else
        (let [href    (escape-html (resolve-link-href node))
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
        (let [href    (resolve-link-href node)
              display (if has-desc
                        (render-inline children :md)
                        (case link-type
                          (:heading :custom-id) target
                          url))]
          (str "[" display "](" href ")")))

      ;; Org format
      :else
      (let [href    (resolve-link-href node)
            display (when has-desc (render-inline children :org))]
        (if display
          (str "[[" href "][" display "]]")
          (str "[[" href "]]"))))))

;; AST Content Rendering
(defn- render-content-in-node [node render-format]
  (let [fmt #(render-inline % render-format)
        render-children* #(mapv (fn [c] (render-content-in-node c render-format)) %)
        result (case (:type node)
                 :document (-> node (update :title #(when % (fmt %))) (update :children render-children*))
                 :section (-> node (update :title #(when % (fmt %))) (update :children render-children*))
                 :paragraph (update node :content fmt)
                 :list (update node :items #(mapv (fn [i] (render-content-in-node i render-format)) %))
                 :list-item (let [node (update node :children render-children*)]
                              (if (:term node)
                                (-> node (update :term fmt) (update :definition fmt))
                                (update node :content fmt)))
                 :table (update node :rows #(mapv (fn [row] (mapv fmt row)) %))
                 :quote-block (update node :children render-children*)
                 :footnote-def (update node :content fmt)
                 :block (if (#{:src :example :export} (:block-type node)) node
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
code { font-family: monospace; }
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

(def ^:private hljs-cdn "https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0")

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

(defn- render-table [rows has-header fmt]
  (if (empty? rows) ""
    (let [;; Find max number of columns across all rows
          max-cols (reduce max 0 (map count rows))
          ;; Pad helper for ragged rows
          pad-row (fn [row] (let [missing (- max-cols (count row))]
                              (if (pos? missing)
                                (into (vec row) (repeat missing []))
                                (vec row))))
          ;; Render all cells to strings, then compute widths from rendered text
          rendered-rows (mapv (fn [row] (mapv #(render-inline % fmt) (pad-row row))) rows)
          col-widths (when (and (seq rendered-rows) (pos? max-cols))
                       (apply mapv (fn [& cells]
                                     (apply max min-table-cell-width (map count cells)))
                              rendered-rows))
          pad-cell (fn [cell width] (str cell (repeat-str (- width (count cell)) " ")))
          format-row (fn [row]
                       (str "| " (str/join " | " (map-indexed #(pad-cell %2 (nth col-widths %1 min-table-cell-width)) row)) " |"))
          separator (when (seq col-widths)
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

(defn- render-list-item [item index ordered level fmt]
  (let [indent (repeat-str (* level list-indent-width) " ")
        marker (if ordered (str (inc index) ". ") "- ")
        children-str (when (seq (:children item))
                       (str/join "\n" (map #(render-node % fmt (inc level)) (:children item))))]
    (if (and (:term item) (= fmt :md))
      (str indent (render-inline (:term item) :md) "\n"
           indent ":   " (render-inline (or (:definition item) []) :md)
           (when children-str (str "\n" children-str)))
      (let [content (if (:term item)
                      (str (render-inline (:term item) fmt) " :: "
                           (render-inline (or (:definition item) []) fmt))
                      (render-inline (:content item) fmt))]
        (str indent marker content (when children-str (str "\n" children-str)))))))

(defn- render-children
  "Render a sequence of child nodes with smart spacing.
   Sections handle their own spacing via :blank-lines-before."
  [children fmt]
  (let [entries (for [child children
                      :let [r (render-node child fmt)]
                      :when (not (str/blank? r))]
                  {:type (:type child) :text r})]
    (->> entries
         (reduce (fn [acc {:keys [type text]}]
                   (if (empty? acc)
                     text
                     (str acc
                          (cond
                            (= type :section) "\n"
                            (= fmt :html)     "\n"
                            :else              "\n\n")
                          text)))
                 ""))))

;; #+OPTIONS parsing
(defn- parse-options-string
  "Parse an #+OPTIONS: value string like 'toc:t H:2 num:t' into a map.
   Values 't' become true, 'nil' become false, integers are parsed,
   other values remain as strings."
  [s]
  (when (and s (not (str/blank? s)))
    (into {}
          (for [[_ k v] (re-seq #"(\S+):(\S+)" s)]
            [(keyword (str/lower-case k))
             (case v
               "t" true
               "nil" false
               (try (Integer/parseInt v) (catch Exception _ v)))]))))

(defn- get-export-options
  "Extract parsed export options from the AST metadata."
  [ast]
  (parse-options-string (get-in ast [:meta :options])))

(defn- collect-toc-entries
  "Recursively collect TOC entries from section nodes up to max-depth.
   Returns a flat list of {:level :title :section-number :id} maps."
  [children max-depth]
  (reduce
   (fn [entries child]
     (if (and (= (:type child) :section)
              (<= (:level child) max-depth))
       (let [entry {:level (:level child)
                    :title (:title child)
                    :section-number (:section-number child)
                    :id (or (get-in child [:properties :custom_id])
                            (heading-to-slug (:title child)))}]
         (into (conj entries entry)
               (collect-toc-entries (:children child) max-depth)))
       entries))
   []
   children))

(defn- render-toc
  "Render a table of contents from TOC entries in the given format."
  [entries fmt]
  (when (seq entries)
    (case fmt
      :html (let [sb (StringBuilder.)
                  _ (.append sb "<nav id=\"table-of-contents\">\n<h2>Table of Contents</h2>\n")
                  _ (loop [[entry & more] entries
                           prev-level 0]
                      (if (nil? entry)
                        ;; Close all remaining open lists
                        (dotimes [_ prev-level]
                          (.append sb "</li>\n</ul>\n"))
                        (let [{:keys [level title section-number id]} entry
                              num-str (if section-number (str section-number " ") "")
                              link (str "<a href=\"#" id "\">" num-str (render-inline title :html) "</a>")]
                          (cond
                            ;; Deeper level: open new sublist(s)
                            (> level prev-level)
                            (do (dotimes [_ (- level prev-level)]
                                  (.append sb "<ul>\n"))
                                (.append sb (str "<li>" link))
                                (recur more level))
                            ;; Same level: close previous item, add new one
                            (= level prev-level)
                            (do (.append sb "</li>\n")
                                (.append sb (str "<li>" link))
                                (recur more level))
                            ;; Shallower level: close sublists, then add item
                            :else
                            (do (dotimes [_ (- prev-level level)]
                                  (.append sb "</li>\n</ul>\n"))
                                (.append sb "</li>\n")
                                (.append sb (str "<li>" link))
                                (recur more level))))))
                  _ (.append sb "</nav>")]
              (.toString sb))
      ;; org and markdown
      (->> entries
           (map (fn [{:keys [level title section-number]}]
                  (let [indent (repeat-str (* 2 (dec level)) " ")
                        num-str (when section-number (str section-number " "))
                        label (if (= fmt :org)
                                (str num-str (render-inline title :org))
                                (str "[" num-str (render-inline title :md) "](#" (heading-to-slug title) ")"))]
                    (str indent "- " label))))
           (str/join "\n")))))

(defn- render-document [node fmt level]
  (let [options (parse-options-string (get-in node [:meta :options]))
        toc? (get options :toc)
        toc-depth (let [h (get options :h 6)
                        t (get options :toc)]
                    ;; toc can be true (use H:) or an integer depth
                    (if (integer? t) t (if (integer? h) h 6)))
        toc-entries (when toc? (collect-toc-entries (:children node) toc-depth))
        toc-str (when toc? (render-toc toc-entries fmt))]
    (case fmt
      :html (let [footnotes (filter #(= (:type %) :footnote-def) (:children node))
                  title-html (when-let [t (:title node)]
                               (str "<header>\n<h1 class=\"document-title\">" (render-inline t :html) "</h1>\n"
                                    (when-let [st (get-in node [:meta :subtitle])]
                                      (str "<p class=\"subtitle\">" (render-inline (parse-inline st) :html) "</p>\n"))
                                    "</header>\n"))
                  toc-html (when toc-str (str toc-str "\n"))
                  main-content (str title-html toc-html
                                    (->> (:children node)
                                         (remove #(= (:type %) :footnote-def))
                                         (map #(render-node % fmt level))
                                         (str/join "\n")))
                  footnotes-html (when (seq footnotes)
                                   (str "<aside class=\"footnotes\">\n"
                                        (str/join "\n" (map #(render-node % fmt level) footnotes))
                                        "\n</aside>"))]
              (html-template (or (inline-text (:title node)) "Untitled Document")
                             (str main-content (when footnotes-html (str "\n" footnotes-html)))))
      :org (let [meta (:meta node)
                 meta-str (cond
                            (seq (:_raw meta)) (str/join "\n" (:_raw meta))
                            (seq (:_order meta))
                            (str/join "\n"
                                      (keep (fn [k]
                                              (let [v (get meta k)]
                                                (when (non-blank? (str v))
                                                  (if (vector? v)
                                                    (str/join "\n" (map #(str "#+" (upper-name k) ": " %) v))
                                                    (str "#+" (upper-name k) ": " v)))))
                                            (:_order meta)))
                            :else nil)]
             (str (when (seq meta-str) (str meta-str "\n\n"))
                  (when toc-str (str toc-str "\n\n"))
                  (render-children (:children node) fmt)))
      ;; default: markdown
      (let [title (if-let [t (:title node)] (str "# " (render-inline t :md) "\n\n") "")]
        (str title
             (when toc-str (str toc-str "\n\n"))
             (render-children (:children node) fmt))))))

(def ^:private planning-order
  "Canonical order for planning keywords."
  [:closed :deadline :scheduled])

(defn- planning-repeaters
  "Build repeater map from a planning map."
  [planning]
  (zipmap planning-order
          (map #(get planning (keyword (str (name %) "-repeat"))) planning-order)))

(defn- planning-items
  "Return [[kw iso] ...] pairs in canonical order, filtering nil values."
  [planning]
  (keep (fn [kw] (when-let [iso (get planning kw)] [kw iso])) planning-order))

(defn- render-planning
  "Render planning info (CLOSED/SCHEDULED/DEADLINE) in the given format."
  [planning fmt]
  (when (seq planning)
    (let [repeaters (planning-repeaters planning)
          items (planning-items planning)]
      (when (seq items)
        (case fmt
          :html (str "<div class=\"planning\">"
                     (str/join " "
                               (map (fn [[kw iso]]
                                      (let [rep (get repeaters kw)
                                            datetime (first (str/split iso #"/"))
                                            display (cond-> (str/replace iso "/" "–")
                                                      rep (str " " rep))]
                                        (str "<span class=\"planning-keyword\">"
                                             (str/upper-case (name kw))
                                             ":</span> <time datetime=\"" datetime "\">" display "</time>")))
                                    items))
                     "</div>")
          :org (str/join " "
                         (map (fn [[kw iso]]
                                (let [rep (get repeaters kw)
                                      rep-str (if rep (str " " rep) "")
                                      ts (if (str/includes? iso "/")
                                           (let [[start end] (str/split iso #"/")
                                                 [d t1] (str/split start #"T")
                                                 t2 (second (str/split end #"T"))]
                                             (str "<" d " " t1 "-" t2 rep-str ">"))
                                           (if (str/includes? iso "T")
                                             (let [[d t] (str/split iso #"T")]
                                               (str "<" d " " t rep-str ">"))
                                             (str "<" iso rep-str ">")))]
                                  (str (str/upper-case (name kw)) ": " ts)))
                              items))
          ;; markdown
          (str/join " "
                    (map (fn [[kw iso]]
                           (let [rep (get repeaters kw)
                                 display (if rep (str iso " " rep) iso)]
                             (str "**" (str/upper-case (name kw)) ":** " display)))
                         items)))))))

(defn- render-section [node fmt]
  (let [planning-str (render-planning (:planning node) fmt)]
    (case fmt
      :html (let [lvl (min (:level node) max-heading-level)
                  tag (str "h" lvl)
                  section-id (or (get-in node [:properties :custom_id])
                                 (heading-to-slug (:title node)))
                  sec-num-html (when-let [sn (:section-number node)]
                                 (str "<span class=\"section-number\">" sn "</span> "))
                  todo-html (when (:todo node)
                              (str "<span class=\"todo " (str/lower-case (name (:todo node))) "\">"
                                   (name (:todo node)) "</span> "))
                  priority-html (when (:priority node)
                                  (str "<span class=\"priority\">[#" (:priority node) "]</span> "))
                  tags-html (when (seq (:tags node))
                              (str " <span class=\"tags\">:" (str/join ":" (:tags node)) ":</span>"))]
              (str "<section id=\"" section-id "\">\n<" tag ">"
                   sec-num-html todo-html priority-html (render-inline (:title node) :html) tags-html
                   "</" tag ">\n"
                   (when planning-str (str planning-str "\n"))
                   (render-children (:children node) fmt) "\n</section>"))
      :org (let [blanks-before (repeat-str (or (:blank-lines-before node) 0) "\n")
                 blanks-after-title (repeat-str (or (:blank-lines-after-title node) 0) "\n")
                 stars (repeat-str (:level node) "*")
                 todo-str (when (:todo node) (str (name (:todo node)) " "))
                 priority-str (when (:priority node) (str "[#" (:priority node) "] "))
                 sec-num-str (when-let [sn (:section-number node)] (str sn " "))
                 tags-str (when (seq (:tags node)) (str " :" (str/join ":" (:tags node)) ":"))
                 props-str (render-properties-org (:properties node))
                 children-str (render-children (:children node) fmt)]
             (str blanks-before
                  stars " " todo-str priority-str sec-num-str (render-inline (:title node) :org) tags-str
                  (when planning-str (str "\n" planning-str))
                  (when props-str (str "\n" props-str))
                  blanks-after-title
                  (when (seq children-str) (str "\n" children-str))))
      ;; default: markdown
      (let [blanks-before (repeat-str (or (:blank-lines-before node) 0) "\n")
            blanks-after-title (repeat-str (or (:blank-lines-after-title node) 0) "\n")
            todo-str (when (:todo node) (str "**" (name (:todo node)) "** "))
            priority-str (when (:priority node) (str "[#" (:priority node) "] "))
            sec-num-str (when-let [sn (:section-number node)] (str sn " "))
            tags-str (when (seq (:tags node)) (str " `:" (str/join ":" (:tags node)) ":`"))
            heading (str (repeat-str (min (:level node) max-heading-level) "#") " "
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
      :org (let [affiliated (:affiliated node)
                 content-str (render-inline content :org)
                 attr-lines (when affiliated
                              (str/join "\n"
                                        (concat
                                         (when-let [caption (:caption affiliated)]
                                           [(str "#+caption: " caption)])
                                         (when-let [aff-name (:name affiliated)]
                                           [(str "#+name: " aff-name)])
                                         (for [[attr-type attrs] (:attr affiliated)
                                               :when (seq attrs)]
                                           (str "#+attr_" (name attr-type) ": "
                                                (str/join " " (for [[k v] attrs]
                                                                (if (str/includes? v " ")
                                                                  (str ":" (name k) " \"" v "\"")
                                                                  (str ":" (name k) " " v)))))))))]
             (if (seq attr-lines)
               (str attr-lines "\n" content-str)
               content-str))
      (render-inline content :md))))

(defn- render-list-node [node fmt level]
  (case fmt
    :html (let [tag (cond (:description node) "dl"
                          (:ordered node) "ol"
                          :else "ul")
                items-html (->> (:items node)
                                (map #(render-node % fmt))
                                (str/join "\n"))]
            (str "<" tag ">\n" items-html "\n</" tag ">"))
    (let [sep (if (and (= fmt :md) (:description node)) "\n\n" "\n")]
      (->> (:items node)
           (map-indexed (fn [idx item] (render-list-item item idx (:ordered node) level fmt)))
           (str/join sep)))))

(defn- render-list-item-node [node fmt level]
  (if (= fmt :html)
    (let [children-html (when (seq (:children node))
                          (str "\n" (str/join "\n" (map #(render-node % fmt) (:children node)))))]
      (if (:term node)
        (str "<dt>" (render-inline (:term node) :html) "</dt>\n<dd>"
             (render-inline (:definition node) :html) children-html "</dd>")
        (str "<li>" (render-inline (:content node) :html) children-html "</li>")))
    (render-list-item node 0 false level fmt)))

(defn- render-table-node [node fmt]
  (if (= fmt :html)
    (let [rows (:rows node)
          has-header (:has-header node)
          cell (fn [tag content] (str "<" tag ">" (render-inline content :html) "</" tag ">"))]
      (if (empty? rows) ""
        (str "<table>\n"
             (when has-header
               (str "  <thead>\n    <tr>\n      "
                    (str/join "" (map #(cell "th" %) (first rows)))
                    "\n    </tr>\n  </thead>\n"))
             "  <tbody>\n"
             (str/join "\n"
                       (map (fn [row]
                              (str "    <tr>\n      "
                                   (str/join "" (map #(cell "td" %) row))
                                   "\n    </tr>"))
                            (if has-header (rest rows) rows)))
             "\n  </tbody>\n</table>")))
    (render-table (:rows node) (:has-header node) fmt)))

(defn- render-src-block [node fmt]
  (case fmt
    :html (str "<pre><code"
               (when (non-blank? (:language node))
                 (str " class=\"language-" (escape-html (:language node)) "\""))
               ">" (escape-html (:content node)) "</code></pre>")
    :org (let [lang (:language node)
               args (:args node)
               escaped-content (escape-block-content (:content node))]
           (str "#+BEGIN_SRC"
                (when (non-blank? lang) (str " " lang))
                (when (non-blank? args) (str " " args))
                "\n" escaped-content "\n#+END_SRC"))
    (str "```" (or (:language node) "") "\n" (:content node) "\n```")))

(defn- render-quote-block [node fmt]
  (case fmt
    :html (str "<blockquote>\n"
               (str/join "\n" (map #(render-node % fmt) (:children node)))
               "\n</blockquote>")
    :org (str "#+BEGIN_QUOTE\n"
              (escape-block-content
               (str/join "\n" (map #(render-inline (:content %) :org) (:children node))))
              "\n#+END_QUOTE")
    (->> (:children node)
         (mapcat (fn [child]
                   (let [text (render-inline (:content child) :md)]
                     (map #(str "> " %) (str/split-lines text)))))
         (str/join "\n"))))

(defn- render-comment-node [node fmt]
  (case fmt
    :html (str "<!-- " (escape-html (:content node)) " -->")
    :org (prefix-lines "# " (:content node))
    (str "<!-- " (escape-html (:content node)) " -->")))

(defn- render-fixed-width [node fmt]
  (case fmt
    :html (str "<pre>" (escape-html (:content node)) "</pre>")
    :org (prefix-lines ": " (:content node))
    (str "```\n" (:content node) "\n```")))

(defn- render-footnote-def-node [node fmt]
  (case fmt
    :html (let [label (escape-html (:label node))]
            (str "<div class=\"footnote\" id=\"fn-" label "\">"
                 "<a href=\"#fnref-" label "\" class=\"footnote-backref\"><sup>" label "</sup></a> "
                 (render-inline (:content node) :html) "</div>"))
    :org (str "[fn:" (:label node) "] " (render-inline (:content node) :org))
    (str "[^" (:label node) "]: " (render-inline (:content node) :md))))

(defn- render-generic-block [node fmt]
  (let [block-type (:block-type node)
        export-type (:args node)]
    (case fmt
      :html (case block-type
              :warning (str "<div class=\"warning\">" (render-inline (parse-inline (:content node)) :html) "</div>")
              :note (str "<div class=\"note\">" (render-inline (parse-inline (:content node)) :html) "</div>")
              :example (str "<pre>" (escape-html (:content node)) "</pre>")
              :export (if (= export-type "html") (:content node) "")
              (str "<div class=\"block-" (name block-type) "\">"
                   "<pre>" (escape-html (:content node)) "</pre></div>"))
      :org (case block-type
             :export (if (= export-type "org")
                       (:content node)
                       (str "#+BEGIN_EXPORT"
                            (when (non-blank? export-type) (str " " export-type))
                            "\n" (escape-block-content (:content node)) "\n#+END_EXPORT"))
             (let [args (:args node)
                   escaped-content (escape-block-content (:content node))]
               (str "#+BEGIN_" (upper-name block-type)
                    (when (non-blank? args) (str " " args))
                    "\n" escaped-content "\n#+END_" (upper-name block-type))))
      ;; markdown
      (case block-type
        :warning (str "> **Warning**\n"
                      (->> (str/split-lines (:content node))
                           (map #(str "> " %))
                           (str/join "\n")))
        :note (str "> **Note**\n"
                   (->> (str/split-lines (:content node))
                        (map #(str "> " %))
                        (str/join "\n")))
        :export (if (= export-type "markdown") (:content node) "")
        :example (str "```\n" (:content node) "\n```")
        (str "```" (name block-type) "\n" (:content node) "\n```")))))

(defn- render-html-line [node fmt]
  (case fmt
    :html (:content node)
    :org (prefix-lines "#+html: " (:content node))
    ""))

(defn- render-latex-line [node fmt]
  (case fmt
    :org (prefix-lines "#+latex: " (:content node))
    ""))

(defn- render-node
  ([node fmt] (render-node node fmt 0))
  ([node fmt level]
   (case (:type node)
     :document       (render-document node fmt level)
     :section        (render-section node fmt)
     :paragraph      (render-paragraph node fmt)
     :list           (render-list-node node fmt level)
     :list-item      (render-list-item-node node fmt level)
     :table          (render-table-node node fmt)
     :src-block      (render-src-block node fmt)
     :quote-block    (render-quote-block node fmt)
     :property-drawer (if (= fmt :org) (render-properties-org (:properties node)) "")
     :comment        (render-comment-node node fmt)
     :fixed-width    (render-fixed-width node fmt)
     :footnote-def   (render-footnote-def-node node fmt)
     :block          (render-generic-block node fmt)
     :html-line      (render-html-line node fmt)
     :latex-line     (render-latex-line node fmt)
     ;; Default case for warnings or unknown types
     (if (:warning node)
       (str "<!-- Warning at line " (:error-line node) ": " (:warning node) " -->")
       ""))))

;; Section numbering
(defn- number-sections
  "Walk the AST and annotate each :section node with a :section-number string
   (e.g. '1', '1.1', '2.3.1') based on its level and position among siblings.
   Only applied when the document has num:t in #+OPTIONS (default: false)."
  [ast]
  (let [options (get-export-options ast)
        num? (get options :num false)]
    (if-not num?
      ast
      (letfn [(number-children [children counters]
                (loop [[child & more] children
                       counters counters
                       result []]
                  (if (nil? child)
                    result
                    (if (= (:type child) :section)
                      (let [level (:level child)
                            ;; Increment counter at this level, reset deeper levels
                            incremented (update counters level (fnil inc 0))
                            updated (apply dissoc incremented (filter #(> % level) (keys incremented)))
                            ;; Build section number string from level 1 to current level
                            sec-num (str/join "." (map #(get updated % 1)
                                                       (range 1 (inc level))))
                            ;; Recursively number children of this section
                            numbered-kids (number-children (:children child) updated)
                            numbered-child (assoc child
                                                  :section-number sec-num
                                                  :children numbered-kids)]
                        (recur more updated (conj result numbered-child)))
                      (recur more counters (conj result child))))))]
        (assoc ast :children (number-children (:children ast) {}))))))

;; ICS Export
(defn- ics-fold-line
  "Fold a content line per RFC 5545 (max 75 octets per line)."
  [line]
  (if (<= (count (.getBytes line "UTF-8")) 75)
    line
    (loop [remaining line first? true parts []]
      (if (empty? remaining)
        (str/join "\r\n " parts)
        (let [max-bytes (if first? 75 74) ;; continuation lines: space takes 1 byte
              ;; Take chars one by one until we hit the byte limit
              chunk (loop [i 1]
                      (if (> i (count remaining))
                        remaining
                        (let [s (subs remaining 0 i)]
                          (if (> (count (.getBytes s "UTF-8")) max-bytes)
                            (subs remaining 0 (dec i))
                            (recur (inc i))))))]
          (recur (subs remaining (count chunk)) false (conj parts chunk)))))))

(defn- ics-escape
  "Escape text for ICS property values per RFC 5545."
  [s]
  (-> s
      (str/replace "\\" "\\\\")
      (str/replace "," "\\,")
      (str/replace ";" "\\;")
      (str/replace "\n" "\\n")))

(defn- iso-to-ics-datetime
  "Convert ISO timestamp (2025-01-15T10:30) to ICS datetime (20250115T103000).
   For intervals (2025-01-15T10:30/2025-01-15T12:00), returns [dtstart dtend].
   For date-only (2025-01-15), returns VALUE=DATE format (20250115)."
  [iso]
  (let [to-ics (fn [s]
                 (if (str/includes? s "T")
                   (let [[d t] (str/split s #"T")
                         date (str/replace d "-" "")
                         digits (str/replace t ":" "")
                         time (cond-> digits (= (count digits) 4) (str "00"))]
                     (str date "T" time))
                   (str/replace s "-" "")))]
    (if (str/includes? iso "/")
      (let [[start end] (str/split iso #"/")]
        {:dtstart (to-ics start) :dtend (to-ics end)})
      {:dtstart (to-ics iso)})))

(defn- collect-ics-items
  "Walk the AST and collect sections that have :scheduled or :deadline planning."
  [node]
  (letfn [(walk [n path]
            (let [is-section (= (:type n) :section)
                  title-text (when is-section (or (inline-text (:title n)) ""))
                  new-path (if is-section (conj path title-text) path)
                  self (when is-section
                         (let [{:keys [planning todo]} n
                               title title-text]
                           (cond-> []
                             (:scheduled planning)
                             (conj {:ics-type :vevent :title title :path new-path :todo todo
                                    :timestamp (:scheduled planning) :repeater (:scheduled-repeat planning)})
                             (and (:deadline planning) todo)
                             (conj {:ics-type :vtodo :title title :path new-path :todo todo
                                    :timestamp (:deadline planning) :repeater (:deadline-repeat planning)}))))]
              (into (vec self) (mapcat #(walk % new-path) (:children n)))))]
    (walk node [])))

(defn- uid-for-item
  "Generate a deterministic UID from item properties."
  [item index]
  (let [input (str (:title item) (:timestamp item) index)
        bytes (.digest (java.security.MessageDigest/getInstance "SHA-256")
                       (.getBytes input "UTF-8"))
        hex (str/join (map #(format "%02x" %) bytes))]
    (str (subs hex 0 16) "@hop")))

(defn- org-repeater-to-rrule
  "Convert Org repeater cookie to ICS RRULE string.
   +1d  -> FREQ=DAILY;INTERVAL=1
   +2w  -> FREQ=WEEKLY;INTERVAL=2
   +1m  -> FREQ=MONTHLY;INTERVAL=1
   +1y  -> FREQ=YEARLY;INTERVAL=1
   .+1d -> FREQ=DAILY;INTERVAL=1  (shift type ignored in ICS)
   ++1w -> FREQ=WEEKLY;INTERVAL=1 (catch-up type ignored in ICS)"
  [repeater]
  (when repeater
    (when-let [[_ interval unit] (re-find #"(\d+)([hdwmy])$" repeater)]
      (let [freq (case unit
                   "h" "HOURLY"
                   "d" "DAILY"
                   "w" "WEEKLY"
                   "m" "MONTHLY"
                   "y" "YEARLY")]
        (str "RRULE:FREQ=" freq ";INTERVAL=" interval)))))

(defn- ics-next-day
  "Given an ICS VALUE=DATE string like '20250115', return the next day '20250116'.
   RFC 5545 requires DTEND to be exclusive for VALUE=DATE events."
  [date-str]
  (-> (java.time.LocalDate/parse date-str
        (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd"))
      (.plusDays 1)
      (.format (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd"))))

(defn- render-ast-as-ics
  "Export scheduled items as VEVENT and deadline+TODO items as VTODO."
  [ast]
  (let [items (collect-ics-items ast)
        now (let [fmt (java.time.format.DateTimeFormatter/ofPattern "yyyyMMdd'T'HHmmss")]
              (.format (java.time.LocalDateTime/now) fmt))
        components
        (map-indexed
         (fn [idx item]
           (let [{:keys [ics-type title todo timestamp repeater]} item
                 uid (uid-for-item item idx)
                 ts (iso-to-ics-datetime timestamp)
                 dtstart (:dtstart ts)
                 is-date (not (str/includes? dtstart "T"))
                 rrule (org-repeater-to-rrule repeater)
                 summary (if (and todo (= ics-type :vtodo))
                           (str (name todo) " " title)
                           title)]
             (case ics-type
               :vevent
               (str/join "\r\n"
                         (concat
                          ["BEGIN:VEVENT"
                           (ics-fold-line (str "UID:" uid))
                           (str "DTSTAMP:" now)]
                          (if is-date
                            [(str "DTSTART;VALUE=DATE:" dtstart)
                             (str "DTEND;VALUE=DATE:" (ics-next-day dtstart))]
                            (if-let [dtend (:dtend ts)]
                              [(str "DTSTART:" dtstart)
                               (str "DTEND:" dtend)]
                              [(str "DTSTART:" dtstart)
                               (str "DTEND:" dtstart)]))
                          (when rrule [rrule])
                          [(ics-fold-line (str "SUMMARY:" (ics-escape summary)))
                           "END:VEVENT"]))

               :vtodo
               (str/join "\r\n"
                         (concat
                          ["BEGIN:VTODO"
                           (ics-fold-line (str "UID:" uid))
                           (str "DTSTAMP:" now)]
                          (if is-date
                            [(str "DUE;VALUE=DATE:" dtstart)]
                            [(str "DUE:" dtstart)])
                          (when rrule [rrule])
                          [(ics-fold-line (str "SUMMARY:" (ics-escape summary)))
                           (str "STATUS:" (if (= todo :DONE) "COMPLETED" "NEEDS-ACTION"))
                           "END:VTODO"])))))
         items)]
    (str (str/join "\r\n"
                   (concat
                    ["BEGIN:VCALENDAR"
                     "VERSION:2.0"
                     "PRODID:-//hop//EN"]
                    components
                    ["END:VCALENDAR"]))
         "\r\n")))

(defn render-ast-as-markdown [ast] (render-node (number-sections ast) :md))
(defn render-ast-as-html [ast] (render-node (number-sections ast) :html))
(defn render-ast-as-org [ast] (render-node (number-sections ast) :org))

;; Output Formatting
(defn format-ast-as-json [ast] (json/generate-string ast {:pretty true}))
(defn format-ast-as-yaml [ast]
  (yaml/generate-string ast {:dumper-options {:flow-style :block}}))

;; Statistics
(defn- count-words
  "Count words in a string or vector of inline nodes."
  [s]
  (let [text (cond (string? s) s
                   (sequential? s) (inline-text s)
                   :else nil)]
    (if (non-blank? text) (count (re-seq #"\S+" text)) 0)))

(defn- count-images
  "Count image links in a vector of inline nodes (or string for backward compat)."
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
  {:words (count-words (:content node))
   :images (count-images (:content node))})

(defn- inc-stat
  "Increment a stat counter in a map, initializing to 0 if absent."
  [m k]
  (update m k (fnil inc 0)))

(defn- add-stat
  "Add n to a stat counter in a map, initializing to 0 if absent."
  [m k n]
  (update m k (fnil + 0) n))

(defn- collect-stats
  "Recursively collect statistics from AST nodes."
  [node]
  (let [node-type (:type node)
        children (:children node [])
        items (:items node [])
        child-stats (reduce (fn [acc child]
                              (merge-with + acc (collect-stats child)))
                            {}
                            (concat children items))
        base-stats (case node-type
                     :section (let [planning (:planning node)]
                                (cond-> (-> child-stats
                                            (inc-stat :sections)
                                            (add-stat :words (count-words (:title node))))
                                  (:scheduled planning) (inc-stat :scheduled)
                                  (:deadline planning)  (inc-stat :deadline)
                                  (:closed planning)    (inc-stat :closed)))
                     (:paragraph :quote-block :footnote-def)
                     (let [{:keys [words images]} (content-stats node)
                           stat-key (case node-type
                                      :paragraph :paragraphs
                                      :quote-block :quote-blocks
                                      :footnote-def :footnotes)]
                       (-> child-stats
                           (inc-stat stat-key)
                           (add-stat :words words)
                           (add-stat :images images)))
                     :table (inc-stat child-stats :tables)
                     :list (inc-stat child-stats :lists)
                     :list-item (-> child-stats
                                    (inc-stat :list-items)
                                    (add-stat :words (+ (count-words (:content node))
                                                        (count-words (:term node))
                                                        (count-words (:definition node))))
                                    (add-stat :images (+ (count-images (:content node))
                                                         (count-images (:definition node)))))
                     :src-block (inc-stat child-stats :src-blocks)
                     :block (inc-stat child-stats :blocks)
                     :comment (inc-stat child-stats :comments)
                     :fixed-width (inc-stat child-stats :fixed-width)
                     :html-line (inc-stat child-stats :html-lines)
                     :latex-line (inc-stat child-stats :latex-lines)
                     :property-drawer (inc-stat child-stats :property-drawers)
                     :document child-stats
                     child-stats)]
    base-stats))

(defn compute-stats
  "Compute statistics for an AST and return a sorted map."
  [ast]
  (let [raw-stats   (collect-stats ast)
        ordered-keys [:sections :paragraphs :words :images :lists :list-items
                      :tables :src-blocks :quote-blocks :blocks
                      :footnotes :comments :fixed-width :html-lines
                      :latex-lines :property-drawers
                      :scheduled :deadline :closed]
        index-map   (zipmap ordered-keys (range))
        present-stats (into {} (filter (fn [[_ v]] (and v (pos? v))) raw-stats))]
    (into (sorted-map-by (fn [a b]
                           (let [c (compare (get index-map a 999)
                                            (get index-map b 999))]
                             (if (zero? c) (compare a b) c))))
          present-stats)))

(def ^:private stats-label-map
  {:sections "Sections"
   :paragraphs "Paragraphs"
   :words "Words"
   :images "Images"
   :lists "Lists"
   :list-items "List items"
   :tables "Tables"
   :src-blocks "Source blocks"
   :quote-blocks "Quote blocks"
   :blocks "Other blocks"
   :footnotes "Footnotes"
   :comments "Comments"
   :fixed-width "Fixed-width blocks"
   :html-lines "HTML lines"
   :latex-lines "LaTeX lines"
   :property-drawers "Property drawers"
   :scheduled "Scheduled"
   :deadline "Deadline"
   :closed "Closed"})

(defn format-stats
  "Format statistics for display."
  [stats]
  (let [max-label-len (->> (keys stats)
                           (map #(count (get stats-label-map % (name %))))
                           (reduce max 0))]
    (->> stats
         (map (fn [[k v]]
                (let [label (get stats-label-map k (name k))
                      padding (repeat-str (- max-label-len (count label)) " ")]
                  (str label ": " padding v))))
         (str/join "\n"))))

;; CLI Options

(defn- resolve-css-theme
  "Resolve a CSS theme value to a map with :link (URL) or :inline (CSS content).
  Resolution order:
  1. https:// URL → {:link url}
  2. file:/// URL → {:inline content} (reads local file)
  3. .css filename matching a local file → {:inline content}
  4. bare name (no spaces) → pico-themes CDN {:link url}"
  [v]
  (cond
    (re-matches #"https?://.*" v)
    {:link v}

    (str/starts-with? v "file:///")
    (let [path (subs v 7)]
      (if (.exists (java.io.File. path))
        {:inline (slurp path)}
        (throw (ex-info (str "CSS theme file not found: " path) {:theme v}))))

    (and (str/ends-with? v ".css") (.exists (java.io.File. v)))
    {:inline (slurp v)}

    (re-matches #"\S+" v)
    {:link (str pico-themes-cdn v ".css")}

    :else
    (throw (ex-info (str "CSS theme not found: " v
                         " (expected a URL, a local .css file, or a pico-theme name)")
                    {:theme v}))))

(def ^:private cli-options
  [["-b" "--base-url URL" "Base URL prepended to relative links (include trailing slash)"]
   ["-c" "--css-theme THEME" "CSS theme: a URL, a local .css file, or a pico-theme name (e.g. org, doric)"]
   ["-f" "--format FORMAT" "Output format: json, edn, yaml, md, html, org, or ics"
    :default "json" :validate [#{"json" "edn" "yaml" "md" "html" "org" "ics"} "Must be: json, edn, yaml, md, html, org, ics"]]
   ["-h" "--help" "Show help"]
   ["-i" "--id REGEX" "Filter: ID or CUSTOM_ID matches" :parse-fn re-pattern]
   ["-I" "--section-id REGEX" "Filter: ancestor ID or CUSTOM_ID matches" :parse-fn re-pattern]
   ["-l" "--level-limit LEVEL" "Filter: level <= LEVEL"
    :parse-fn #(Integer/parseInt %) :validate [pos? "Must be positive"]]
   ["-L" "--level-limit-inclusive LEVEL" "Filter: level <= LEVEL and render deeper headings as bold"
    :parse-fn #(Integer/parseInt %) :validate [pos? "Must be positive"]]
   ["-n" "--no-unwrap" "Preserve original line breaks"]
   ["-r" "--render FORMAT" "Content rendering format in AST output: md, html, or org"
    :default "md" :validate [#{"md" "html" "org"} "Must be: md, html, org"]]
   ["-s" "--stats" "Compute and display document statistics"]
   ["-t" "--title REGEX" "Filter: section title matches" :parse-fn re-pattern]
   ["-T" "--section-title REGEX" "Filter: ancestor title matches" :parse-fn re-pattern]])

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
  ([] (System/exit 0))
  ([msg] (println msg) (System/exit 0)))

(defn -main [& args]
  (let [{:keys [options arguments errors summary]} (cli/parse-opts args cli-options)]
    (cond
      (:help options)
      (exit-ok (usage summary))

      errors
      (exit-error (str/join "; " errors) summary)

      (not= (count arguments) 1)
      (exit-error "Expected one argument (org-file or '-')" summary)

      :else
      (let [file-path (first arguments)
            org-content (if (= file-path "-")
                          (slurp *in*)
                          (try (slurp file-path)
                               (catch java.io.FileNotFoundException _
                                 (exit-error (str "File not found - " file-path)))))]
        (try
          (binding [*base-url* (when-let [u (:base-url options)]
                                 (cond-> u (not (str/ends-with? u "/")) (str "/")))
                    *css-theme* (when-let [v (:css-theme options)]
                                  (resolve-css-theme v))]
            (let [unwrap? (not (:no-unwrap options))
                  ast (organ/parse-org org-content {:unwrap? unwrap?})
                  filter-opts {:level-limit (:level-limit options)
                               :level-limit-inclusive (:level-limit-inclusive options)
                               :title-pattern (:title options)
                               :id-pattern (:id options)
                               :section-title-pattern (:section-title options)
                               :section-id-pattern (:section-id options)}
                  filtered-ast (organ/filter-ast ast filter-opts)
                  output-format (:format options)
                  render-format (keyword (:render options))
                  is-org-output (= output-format "org")
                  cleaned-ast (if is-org-output filtered-ast (organ/clean-node filtered-ast))]
              ;; Report parse errors to stderr if any
              (when-let [errs (:parse-errors cleaned-ast)]
                (binding [*out* *err*]
                  (doseq [{:keys [line message]} errs]
                    (println (str "Warning (line " line "): " message)))))
              (cond
                (:stats options)
                (let [stats (compute-stats filtered-ast)]
                  (println (format-stats stats)))

                (= output-format "md")
                (println (render-ast-as-markdown cleaned-ast))

                (= output-format "html")
                (println (render-ast-as-html cleaned-ast))

                (= output-format "org")
                (println (render-ast-as-org cleaned-ast))

                (= output-format "ics")
                (do (print (render-ast-as-ics filtered-ast))
                    (flush))

                :else
                (let [rendered-ast (render-content-in-node cleaned-ast render-format)]
                  (case output-format
                    "json" (println (format-ast-as-json rendered-ast))
                    "edn" (println (organ/format-ast-as-edn rendered-ast))
                    "yaml" (println (format-ast-as-yaml rendered-ast)))))
              (System/exit 0)))
          (catch clojure.lang.ExceptionInfo e
            (exit-error (str "Parse error: " (.getMessage e))))
          (catch Exception e
            (binding [*out* *err*]
              (println "Error:" (.getMessage e))
              (.printStackTrace e *err*))
            (System/exit 1)))))))

(when (= *file* (System/getProperty "babashka.file"))
  (apply -main *command-line-args*))
