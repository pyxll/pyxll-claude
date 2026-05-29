---
name: fetch-pyxll-docs
description: Fetch the PyXLL documentation and use it as context when writing, reviewing, or debugging PyXLL code. Invoke automatically whenever the task involves xl_func, xl_menu, create_ctp, rebind, ribbon XML, or any other PyXLL-specific API.
user-invocable: false
metadata:
  author: pyxll
---

Fetch the PyXLL documentation and use it as context for the current task.

## Steps

1. Fetch the documentation index:

   ```
   https://www.pyxll.com/llms.txt
   ```

   The index contains a navigation guide at the top explaining which sections to
   read for common tasks, followed by links to individual `.md` pages grouped by
   topic (User Guide, API Reference, Changelog, etc.).

2. Based on the current task, identify the relevant pages from the index and fetch
   each one. Use the navigation guide at the top of the index to find the right
   sections quickly. Fetch all pages relevant to the task — do not skip API
   reference pages when writing code.

3. If the information needed was not found in the pages fetched from the index,
   fall back to the full concatenated docs for a deeper search. Because this file
   is ~500 KB, use Bash with curl rather than the Fetch tool to avoid truncation:

   ```bash
   curl -s https://www.pyxll.com/llms-full.txt
   ```

4. Use the documentation to inform your answer, code, or review.

## Rules

- ALWAYS fetch these docs before writing any PyXLL-specific code.
- Do NOT rely on training-data knowledge alone for PyXLL APIs — the docs are authoritative.
- When writing `@xl_func` functions, check the type signature and argument type syntax from the docs.
- When subclassing a PyXLL class (e.g. Formatter, ConditionalFormatterBase), fetch the API reference for that
  class to get exact method signatures before looking at examples.
