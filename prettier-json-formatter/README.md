# Prettier JSON Formatter

A single offline HTML file for pasting, checking, editing and copying JSON.
No install, no internet, and nothing leaves your computer.

## Use it

Double-click `PrettierJSON.html` to open it in Chrome or Edge, then:

1. Paste JSON with Ctrl+V. It is formatted automatically.
2. Edit it in the editor.
3. Click **Copy** or press Ctrl+Shift+C to copy all of it.

## Features

- **Format** (Ctrl+Shift+F) in Prettier style: one key per line, and short arrays
  of plain values stay on one line if they fit the width.
- **Minify** (Ctrl+Shift+M) and **Sort keys** (A→Z, recursive).
- Indent: 2 spaces, 4 spaces or tabs. Width: 80, 120 or expand all.
- Syntax colors like the Prettier playground: keys, strings, numbers, true/false and
  null each get their own color, and the line with an error is highlighted in red.
  Turn this off with the **Colors** checkbox.
- Undo and redo (Ctrl+Z, Ctrl+Shift+Z or Ctrl+Y), including after Format or Minify.
- Live validation. An error shows its line and column, and you can click it to jump there.
- Numbers and strings are kept exactly as written, so big IDs like
  `12345678901234567890` and values like `1.10` are never rounded or changed.
- **Fix loose JSON**: accepts comments, trailing commas, single quotes, unquoted keys,
  `True`/`False`/`None` and `undefined`, and outputs strict JSON.
- Open or drop a `.json` file, and download the result.
- Dark mode by default (switch to light with the ☀ Light button). Remembers your settings and last text in this browser.

## Sample data

Open these with **Open file**, or drag them into the editor:

| File | Size | What it tests |
|---|---|---|
| `samples/tickets-12k.json` | 12 KB, minified | Format. 18-digit IDs and `0.50`-style numbers must stay exactly as written |
| `samples/tickets-18k-pretty.json` | 18 KB, formatted | Minify and Sort keys |
| `samples/tickets-155k.json` | 155 KB | Large file, scrolling, colors |
| `samples/tickets-780k.json` | 780 KB | Very large file |
| `samples/long-value-29k.json` | 29 KB | Single string values over 10,000 characters |
| `samples/loose.json5` | small | Comments, single quotes, bare keys, trailing commas, `True`/`None`/`undefined` → strict JSON |
| `samples/broken.json` | small | Missing comma: error on line 6 |
| `samples/tickets-155k-broken.json` | 155 KB | Missing comma near the end of a large minified file |
