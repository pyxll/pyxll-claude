---
name: fetch-pyxll-docs
description: Fetch the PyXLL documentation and use it as context when writing, reviewing, or debugging PyXLL code. Invoke automatically whenever the task involves xl_func, xl_menu, create_ctp, rebind, ribbon XML, or any other PyXLL-specific API.
user-invocable: false
metadata:
  author: pyxll
---

Fetch the PyXLL documentation and use it as context for the current task.

## Steps

1. Fetch the documentation index to find relevant pages:

   ```bash
   curl -s https://www.pyxll.com/llms.txt
   ```

   Read the **entire** output without truncating, piping to `head`, or
   summarising. The index contains a navigation guide at the top mapping
   common tasks to sections, followed by page titles, descriptions, and URLs
   grouped by topic. Truncating it will cause you to miss relevant pages.

2. Fetch the individual pages relevant to the task directly by their URL:

   ```bash
   curl -s <page-url>
   ```

3. If the index does not surface what you need, use the search script to find
   pages by keyword. The script caches the full docs locally (refreshed every
   24 h) and returns only matching page URLs — avoiding loading 500 KB into context.

   The script is in the `scripts/` folder next to this file. Find this file's
   location and run:

   ```bash
   /path/to/this/skills/fetch-pyxll-docs/scripts/search-pyxll-docs.sh <keyword> [keyword2 ...]
   ```

   Then fetch the returned page URLs individually using curl.

4. Use the documentation to inform your answer, code, or review.

## Rules

- ALWAYS fetch these docs before writing, modifying, or troubleshooting any
  PyXLL-specific code or behaviour. Before suggesting a manual workaround for a
  PyXLL problem, check whether PyXLL already has a built-in solution (decorator
  parameter, config key, or feature).
- Do NOT rely on training-data knowledge alone for PyXLL APIs — the docs are authoritative.
- When writing `@xl_func` functions, check the type signature and argument type
  syntax from the docs.
- Before writing any code that calls the Excel COM API (Range, Worksheet, Workbook, etc.),
  fetch https://www.pyxll.com/docs/userguide/vba.md and read it in full. It documents
  critical differences between VBA and Python — including how COM properties that take
  arguments must be called as `Get<PropertyName>(args)` in Python rather than
  `Property(args)` as in VBA.
- Before using any PyXLL class, function, decorator, or configuration setting
  (including pyxll.cfg section names and their keys), fetch the relevant documentation
  and use only what is explicitly documented. Never infer behaviour, key names, or
  parameter names from conventions or assumptions — if it is not in the docs, do not
  use it.
