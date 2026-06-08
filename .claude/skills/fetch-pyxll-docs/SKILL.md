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

   The index contains a navigation guide at the top mapping common tasks to
   sections, followed by page titles, descriptions, and URLs grouped by topic.

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

- ALWAYS fetch these docs before writing any PyXLL-specific code.
- Do NOT rely on training-data knowledge alone for PyXLL APIs — the docs are authoritative.
- When writing `@xl_func` functions, check the type signature and argument type syntax from the docs.
- Before using any PyXLL class, function, or decorator, fetch its API reference and 
  use only what is explicitly documented. Never infer behaviour from conventions or 
  assumptions — if it is not in the docs, do not use it.
