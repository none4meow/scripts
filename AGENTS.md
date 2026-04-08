# AGENTS.md

## Purpose

This repository contains Adobe Illustrator ExtendScript (`.jsx`) automation scripts.

The main priorities are:

1. Correct object placement in Illustrator.
2. Safe, minimal edits to existing scripts.
3. Predictable geometry, spacing, and color grouping behavior.
4. Readable ExtendScript that is easy to tweak by hand later.

## File-first rule

- Always read the current file before making changes.
- Do not assume the latest script structure from memory.
- Before editing, inspect the actual file contents and base all changes on that version.
- If asked to modify a script, prefer patching the existing file over rewriting from scratch.
- If the file content is unavailable, say so clearly and ask for the file instead of guessing.

---

## Working style

- Make the smallest change that solves the task.
- Prefer modifying existing logic over rewriting the whole script.
- Keep scripts single-file unless the task clearly requires splitting code.
- Before changing behavior, identify which constants and helper functions control it.
- Preserve existing variable names when possible so the script stays familiar.

---

## Illustrator scripting rules

- Target Adobe Illustrator ExtendScript, not browser JavaScript and not Node.js.
- Assume the runtime is old JavaScript: avoid modern syntax that ExtendScript may not support.
- Use `var`, not `let` or `const`.
- Avoid external dependencies.
- Do not introduce build steps, npm packages, or transpilers unless explicitly requested.
- Prefer simple loops and plain helper functions over clever abstractions.
- Keep coordinate math explicit and commented.

---

## Geometry and placement rules

- Treat bounds carefully. Be explicit about whether logic uses:
  - `visibleBounds`
  - `geometricBounds`
  - union of both
- Do not silently switch bounds behavior. If changing bounds logic, explain why.
- When packing objects, preserve these assumptions unless the user asks otherwise:
  - no rotation
  - bottom-to-right placement order
  - row wrapping when needed
  - box-based packing
- Keep spacing logic centralized in constants such as:
  - box size
  - box padding
  - object gap
  - grid cell size
  - label offset
- If a visual gap is caused by bounding boxes of rounded objects, do not treat it as a bug unless the user says it is.

---

## Color grouping rules

When grouping objects by color:

- Use the first rendered item inside a group as the base layer when the task says “first layer color”.
- Interpret “first layer” as the back-most visible drawable item.
- Prefer fill color first, then stroke color as fallback, unless the task says to use both.
- If using both fill and stroke in a grouping key, keep that behavior explicit in code comments.
- Quantization or tolerance must be defined in one place so it is easy to tune later.

---

## UI rules

If a script includes a menu or dialog:

- Keep the UI minimal.
- Prefer simple ScriptUI dialogs with clear button labels.
- Buttons should directly map to script behavior.
- Do not add unnecessary options.

---

## Editing rules

When updating a script:

- Do not remove working behavior unless asked.
- Do not add heavy or slow algorithms if a simpler fast approach already matches the user’s intent.
- If a new approach is slower or riskier, keep the old logic available or revert cleanly.
- Preserve comments that explain user-facing behavior.

---

## Code quality rules

- Add short comments only where they clarify Illustrator-specific behavior.
- Avoid over-commenting obvious lines.
- Name helpers clearly, for example:
  - `getUnionBounds`
  - `findBackMostDrawable`
  - `moveItemBottomLeftTo`
  - `createColorLabel`
- Keep related constants together at the top of the file.
- Keep helper functions above the main execution block.

---

## Output expectations

When asked to generate or revise a script:

- Return complete runnable `.jsx` code unless the user asks for a partial patch.
- If only a small patch is needed, show exactly what to replace.
- Preserve the current behavior outside the requested change.
- Be explicit about any assumptions.

---

## Review checklist

Before finishing, check:

- Is the code valid for Illustrator ExtendScript?
- Are bounds and coordinate assumptions consistent?
- Are spacing constants applied consistently?
- Are box layout rules consistent?
- Is color grouping based on the requested first-layer logic?
- Did the change stay as small as possible?