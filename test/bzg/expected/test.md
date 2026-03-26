# Test Document

- [1 First Section](#first-section)
  - [1.1 Nested Section](#nested-section)
    - [1.1.1 Deep Section](#deep-section)
- [2 Second Section [#A]](#second-section-a)
- [3 Lists](#lists)
  - [3.1 Unordered List](#unordered-list)
  - [3.2 Ordered List](#ordered-list)
  - [3.3 Description List](#description-list)
- [4 Tables](#tables)
- [5 Blocks](#blocks)
  - [5.1 Source Block](#source-block)
  - [5.2 Quote Block](#quote-block)
  - [5.3 Example Block](#example-block)
- [6 Inline Elements](#inline-elements)
- [7 Fixed Width and Comments](#fixed-width-and-comments)
- [8 Footnotes](#footnotes)

Some top-level paragraph with **bold**, *italic*, _underline_, ~~strikethrough~~, `code`, and `verbatim` markup.

A second paragraph with a [link with description](https://example.com) and a bare link [https://example.com](https://example.com).

# 1 First Section

Paragraph in first section.

## 1.1 Nested Section

Paragraph in nested section with footnote reference[^1].

### 1.1.1 Deep Section

Deep content here.

# **TODO** 2 Second Section [#A] `:tag1:tag2:`
**SCHEDULED:** 2025-03-15T10:00

This section has a TODO keyword, priority, tags, scheduling and properties.

# 3 Lists

## 3.1 Unordered List

- First item
- Second item with **bold**
- Third item
  - Nested item one
  - Nested item two

## 3.2 Ordered List

1. First
2. Second
3. Third

## 3.3 Description List

Emacs
:   A self-documenting text editor

Vim
:   A modal text editor

Nano
:   A simple terminal editor

# 4 Tables

| Name  | Language | Year |
|-------|----------|------|
| Emacs | Lisp     | 1976 |
| Vim   | C        | 1991 |
| Nano  | C        | 1999 |

# 5 Blocks

## 5.1 Source Block

```clojure
(defn hello [name]
  (println "Hello," name))
```

## 5.2 Quote Block

> The best way to predict the future is to invent it.
> -- Alan Kay

## 5.3 Example Block

```
This is example text.
  Indentation is preserved.
```

# 6 Inline Elements

An image link: ![https://example.com/image.png](https://example.com/image.png).

An image with description: [A photo](https://example.com/photo.jpg).

A mailto link: [Email me](mailto:test@example.com).

An internal link to [second-section](#second-section).

# 7 Fixed Width and Comments

```
This is fixed width text.
Second fixed width line.
```

<!-- This is a comment.
Second comment line. -->

# 8 Footnotes

Paragraph referencing another footnote[^2].

[^1]: First footnote definition.

[^2]: Second footnote definition.
