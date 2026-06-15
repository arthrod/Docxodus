# IR-Powered DOCX Editor — Roadmap

Companion to `ir_editor_feasibility.md` (which records the verdict, architecture, and
PoC results). This is the **sequenced, prioritized** plan for turning the proven
foundation + MVP into a complete editor. Supersedes the scattered "Still Plan 2" notes.

Status as of PR #234 (branch `feat/ir-editor-feasibility-poc`): **foundation + MVP
shipped and proven.** This roadmap is the work that remains.

## Architecture invariants (do not break)

1. **Model-of-record = the live OOXML in `DocxSession`** (lossless `Save`). The IR is
   read-only and has no IR→OOXML writer; never make the IR or the DOM the source of truth.
2. **Addressing = the shared `{#kind:scope:unid}` anchor system.** `convertDocxToHtml`
   (stampAnchors) ↔ `DocxSession` ↔ `RenderBlock` all use one Unid scheme; keep it that way.
3. **Render is a projection; patch incrementally.** An edit goes through `DocxSession` by
   anchor, then only the changed block re-renders (`RenderBlockHtml`). Never round-trip the
   whole doc through `convertDocxToHtml` per edit.
4. **Untouched content stays byte-faithful on save.** Edits may simplify the *edited* block
   (within the markdown subset), but must never degrade blocks the user didn't touch.

## Shipped (foundation + MVP)

- C#: `WmlToHtmlConverterSettings.StampAnchors`; `HtmlConversionOps.RenderBlockHtml`
  (stateless + session-attached); `DocxSession.LiveDocument`.
- WASM/npm: `RenderBlockHtml`, `stampAnchors`, `renderBlockHtml()`, `DocxSession.renderBlock()`.
- `DocxEditor` (pure TS): faithful render → editable paragraphs/headings → commit via
  `DocxSession` → incremental re-render → lossless save; `{ paginated: true }` page boxes.
- Tests: C# `HCO050`/`HCO052`; browser `render-block.spec.ts`, `editor.spec.ts`.

## Milestones (priority order = impact)

### M1 — Rich in-block editing (preserve + apply inline formatting)  · effort M–L · **next**
**Problem:** `commitBlock` replaces an edited block from `el.textContent` (plain text), so
editing a formatted paragraph **destroys its bold/italic/links** — the biggest correctness
trap.
**Approach:** add an HTML-inline → markdown serializer that walks the edited block's DOM and
emits the projector's markdown subset, detecting emphasis via `getComputedStyle`
(font-weight/font-style), links via `href`, code via the run's style — then
`ReplaceText(anchor, markdown)` instead of plain text. (Finer-grained `ReplaceTextAtSpan` /
`ApplyFormat` is a later optimization; markdown round-trip covers the common cases first.)
**Acceptance:** editing a paragraph that contains a bold/italic/linked run, then save+reopen,
preserves that formatting (projection shows `**…**` / `*…*` / `[…](…)`). Formatting the
markdown subset can't express (size/color) is still dropped on an *edited* block — documented.

### M2 — Structural editing via keyboard  · effort M
**Problem:** no way to add/split/merge/delete blocks from the UI; ops exist in `DocxSession`
but aren't wired.
**Approach:** Enter at a caret → `SplitParagraph(anchor, offset)`; Backspace at block start →
`MergeParagraphs(prev, this)`; a block-insert affordance → `InsertParagraph`; delete-empty →
`DeleteBlock`. Reconcile the DOM from `EditResult.Created/Removed/Modified` (patch the
affected nodes, not a full re-render). Maintain the `unid → fullId` map across these deltas.
**Acceptance:** browser test — split a paragraph, merge two, insert and delete a block; each
reflected incrementally and surviving save/reopen.

### M3 — Worker offload  · effort M–L
**Problem:** the initial full convert (~0.7–2.4 s) and session ops run on the main thread →
the UI freezes on open and on big docs. (Per-edit is already ~10 ms.)
**Approach:** extend the Web Worker surface (`docxodus.worker.ts` / `worker-proxy.ts`) to
carry session open/edit/render-block/save, transfer bytes zero-copy; the main thread holds
only the DOM. Keep the synchronous `DocxEditor` API working by awaiting worker round-trips.
**Acceptance:** opening and editing a large doc never blocks the main thread > ~16 ms;
existing editor tests pass through the worker path.

### M4 — Re-paginate on edit  · effort M
**Problem:** in paginated mode an edited block can overflow its page box (the MVP patches in
place without reflowing).
**Approach:** after a commit in paginated mode, re-run pagination from the affected page
forward (staging originals are retained, so a scoped reflow is feasible); debounce.
**Acceptance:** an edit that grows a block past a page boundary reflows to a new page.

### M5 — Formatting toolbar + undo/redo  · effort S–M
**Approach:** bold/italic/style/list controls → `ApplyFormat`/`SetParagraphStyle`/
`SetListLevel`; Ctrl+Z/Y → `DocxSession.Undo/Redo` (+ re-render affected blocks). Mostly UI
glue over existing ops.
**Acceptance:** toolbar applies formatting to a selection; undo/redo round-trips an edit.

### M6 — Tracked-changes / review mode  · effort M
**Approach:** open the session with `TrackedChanges = RenderInline`; render `ins`/`del` with
author colors; serve the redline/review use case.
**Acceptance:** edits land as `w:ins`/`w:del` with author attribution, visible in the editor.

### M7 — Table-cell & table-structure editing  · effort M
**Approach:** `ReplaceCellContent` for cell text; row/col insert-delete and cell-merge via
`session.Raw.*` until first-class ops exist.
**Acceptance:** a cell's content edits and round-trips; tables are no longer read-only.

### M8 — React wrapper  · effort S
**Approach:** `useDocxEditor` hook + `<DocxEditor>` component over the pure-TS core, in
`npm/src/react.ts`.
**Acceptance:** a React app mounts the editor with one component.

### M9 — Single-block render fidelity  · effort M
**Approach:** copy image parts into the throwaway render doc; resolve list-numbering
continuation for a block rendered in isolation.
**Acceptance:** re-rendering an image-bearing or list-item block matches the full render.

## Recommended sequencing

**M1 + M2 together = "make editing real"** (the difference between a demo and a usable
editor), then **M3** for responsiveness. M4–M9 sequence by target use case: authoring favors
M4/M5/M2; review favors M6; broad fidelity favors M7/M9. M8 (React) any time.
