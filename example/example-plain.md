---
marp: true
theme: default
paginate: true
title: Example deck
---

<!-- _class: lead -->
<!-- _paginate: false -->

# Example deck

A test of **marp-to-editable-pptx-cli**

Every word on these slides should be editable in PowerPoint

---

# Text formatting

Plain text, **bold**, *italic*, ~~strikethrough~~ and `inline code`.

A [link to Marp](https://marp.app/) should stay clickable.

> A blockquote, for a quotation or an aside.

---

# Lists

- First bullet
- Second bullet
  - Nested bullet
  - Another nested bullet
- Third bullet

1. Numbered item
2. Numbered item
3. Numbered item

---

# Table

| Feature  | Marp CLI `--pptx` | This tool |
| -------- | :---------------: | :-------: |
| Text     | Image             | Editable  |
| Lists    | Image             | Editable  |
| Tables   | Image             | Editable  |

---

# Code

```js
const { marpCli } = require('marp-to-editable-pptx-cli')

const status = await marpCli(['example.md', '--force'])
console.log(status === 0 ? 'Success' : 'Failure')
```

---

<!-- _backgroundImage: none -->
<!-- _backgroundColor: #1e3a5f -->
<!-- _color: red -->

# Slide colours

This slide sets its own background and text colour with Marp directives.

<!--
Speaker notes: this comment should become the slide's notes in PowerPoint.
-->

---

# The end

Open `example.pptx` and check that the text, lists, table and code can be edited.
