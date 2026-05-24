# Markdown Projection

> **Status:** Design — scaffold only. The `WmlToMarkdownConverter` class exists in source with the public surface defined, but conversion logic is not yet implemented. This document is the spec the implementation will follow.

The markdown projection is a deterministic, **anchor-addressed** rendering of a DOCX as Markdown. It is a sibling to `WmlToHtmlConverter` and `OpenContractExporter` in the converter family, intended as a substrate for tooling that wants to operate on Word documents the way it would operate on source files — search, splice, diff, address by ID. Use cases include LLM-driven editing pipelines, structured search indexers, and diff/review UIs that need a text view richer than `WmlToHtmlConverter` strips down to.

## Goals

1. **Stable addressing.** Every paragraph, heading, list item, table, table cell, comment, footnote, and endnote in the projection is reachable by an anchor that survives reformatting and reordering.
2. **Deterministic output.** Two runs on the same input produce byte-identical output.
3. **Round-trippable references.** An anchor in the projection maps unambiguously back to an `XElement` (or set of elements) in the underlying OOXML. The projection itself is read-only — mutation lives in callers — but the anchor → element resolver is part of this module's API.
4. **Lossy by design, honestly.** Anything that can't fit in Markdown becomes either a structured opaque anchor (with metadata callers can fetch via the SDK) or a clearly-marked annotation. Silent loss is a bug.

## Non-Goals

- **Round-trip rendering** (Markdown → DOCX). That problem already has `HtmlToWmlConverter` and `DocumentAssembler`; this converter is one-way.
- **GFM-perfect tables.** Word tables (merged cells, nested tables, cell-level shading, vertical text, …) exceed GFM. We render what fits and surface the rest as opaque anchors.
- **Preserving every formatting nuance.** Bold/italic/code/links/headings/lists/quotes carry over. Font sizes, colors, character spacing, etc. do not — they're recoverable by anchor lookup if needed.

## Anchor Scheme

Anchors derive from the `Unid` system Docxodus already maintains on paragraphs and runs (see `AssignUnidToAllElements` / the legacy migration notes in `CLAUDE.md`). They are stable across edits unless the underlying element is removed.

**Format:** `{#kind:scope:unid}` where:

| Field | Values | Meaning |
|---|---|---|
| `kind` | `p`, `h`, `li`, `tbl`, `tr`, `tc`, `cmt`, `fn`, `en`, `img`, `drw`, `unk` | Element type |
| `scope` | `body`, `hdr1`…`hdrN`, `ftr1`…`ftrN`, `fn`, `en`, `cmt` | Which part of the package |
| `unid` | 8–16 hex chars | Stable element identifier |

Examples:

- `{#h:body:a1b2c3d4}` — heading in the main body
- `{#p:hdr1:9f8e7d6c}` — paragraph in the first header part
- `{#tc:body:1a2b3c4d}` — table cell in the body
- `{#cmt:cmt:00ff11ee}` — comment in the comments part

Anchors appear at the start of the line they refer to (block-level) or as inline `{#…}` markers (inline annotations like comments anchored to a span).

### Why prefix-with-`#`?

`{#…}` is the [Pandoc / kramdown attribute syntax](https://pandoc.org/MANUAL.html#extension-header_attributes). Most Markdown renderers either honor it (turning anchors into `id` attributes) or display it literally without breaking layout. Either way it survives copy/paste round trips through agent contexts.

## Element Coverage

| OOXML element | Markdown representation | Notes |
|---|---|---|
| `w:p` (no style) | Paragraph | Anchor on its own line above the text |
| `w:p` styled `Heading{1..6}` | `#`…`######` heading | Heading level taken from style |
| `w:p` styled `Title` / `Subtitle` | `#` heading with class | |
| `w:p` with `w:numPr` | `-` or `1.` list item | Numbering resolved to literal markers; nested via indentation |
| `w:p` styled `Quote` / `IntenseQuote` | `>` blockquote | |
| `w:p` styled `Code` / `HTML Code` | Indented or fenced code block | |
| `w:r` with `w:b` | `**bold**` | |
| `w:r` with `w:i` | `*italic*` | |
| `w:r` with `w:rStyle="Code"` or monospace | `` `code` `` | |
| `w:r` with `w:strike` | `~~strike~~` | GFM extension |
| `w:hyperlink` | `[text](url)` | Internal links use anchor: `[text](#anchor)` |
| `w:tbl` (simple) | GFM pipe table | When no merged cells, nesting, or per-cell formatting |
| `w:tbl` (complex) | Opaque anchor block | `{#tbl:body:…}` followed by a fenced `text` block with a structural summary |
| `w:commentRangeStart`…`End` | Inline `{#cmt:cmt:…}` markers wrapping the commented span | Comment text appears in a Comments section at end |
| `w:footnoteReference` | `[^fn-xxxx]` GFM footnote ref | Definitions collected at end |
| `w:endnoteReference` | `[^en-xxxx]` | Same |
| `w:drawing` / `w:pict` (image) | `![alt](docxodus://img/…){#img:…}` | URL is a scheme the caller resolves; metadata accessible via anchor |
| `w:sdt` (content control) | Rendered content, anchor on outer SDT | The SDT itself is an anchor target so callers can address "this content control" |
| `w:ins` / `w:del` (tracked changes) | Configurable: accept, show as `{+ins+}`/`{-del-}`, or omit | Mirrors `WmlToHtmlConverter.RenderTrackedChanges` |
| `w:sectPr` | `---` thematic break with section anchor | Sections become navigable |

Anything not in the table above renders as a single line:

```
{#unk:body:…} [unsupported: w:smartTag]
```

## Multipart Namespacing

A DOCX has many "documents" — the body is one part, but headers, footers, footnotes, endnotes, and comments live in sibling parts. The projection emits them as named sections so a single Markdown stream covers the whole package:

```markdown
# Document

{#p:body:…} The Provider shall...

---

# Headers

## hdr1
{#p:hdr1:…} CONFIDENTIAL

# Footers

## ftr1
{#p:ftr1:…} Page {PAGE} of {NUMPAGES}

# Footnotes

[^fn-aaaa]: {#fn:fn:aaaa} See Section 4.2 for definitions.

# Comments

- {#cmt:cmt:bbbb} **Alice** (2026-05-23): Should this be capitalized?
```

Callers that only care about the body can pass `Scopes = ProjectionScopes.Body` to skip the rest.

## Numbering Resolution

Word list numbers are computed from `w:numPr` referencing `numbering.xml` (`w:abstractNum` / `w:lvlText`), not stored as text. The projection **resolves numbering** so the agent sees what a human reads:

```markdown
{#li:body:…} 1. First item
{#li:body:…} 2. Second item
{#li:body:…}   a. Nested item
{#li:body:…} 3. Third item
```

The original `numPr` is recoverable through the anchor — callers that want to edit the list's numbering format address the source, not the rendered number.

Legal numbering (`1.1`, `1.1.1`, …) and other multi-level formats render verbatim.

## Tables

GFM pipe tables when:

- No merged cells (`w:gridSpan`, `w:vMerge`)
- No nested tables
- No per-cell formatting beyond bold/italic in cell content
- ≤ ~80 chars per cell (configurable)

Otherwise, an opaque anchor block:

````markdown
{#tbl:body:t1}
```table
rows: 4
cols: 3
caption: Fee Schedule
notes: merged cells in row 1; nested table in (3,2)
```
````

Per-cell content is reachable via `{#tc:body:…}` anchors that the caller can fetch individually with the SDK. The opaque block keeps the projection readable; the anchors keep it addressable.

## Round-Trip: Anchor → XElement

The companion API on the converter:

```csharp
var projection = WmlToMarkdownConverter.Convert(wmlDoc, settings);
// projection.Markdown is the text
// projection.AnchorIndex is an IReadOnlyDictionary<string, AnchorTarget>

var target = projection.AnchorIndex["p:body:a1b2c3d4"];
// target.PartUri, target.ElementXPath, target.Unid
// target.Resolve(WordprocessingDocument) -> XElement
```

This is the contract that makes the projection useful for editing: callers receive the projection *and* a way to walk back to the source for any anchor.

## Settings (Planned)

```csharp
public class WmlToMarkdownConverterSettings
{
    // What parts of the package to include.
    public ProjectionScopes Scopes = ProjectionScopes.All;

    // Heading level offset (e.g., 1 means Word Heading1 -> Markdown ##).
    public int HeadingLevelOffset = 0;

    // Inline anchors? Block-level only? Or omit?
    public AnchorRenderMode AnchorMode = AnchorRenderMode.Block;

    // How to handle complex tables.
    public TableRenderMode TableMode = TableRenderMode.GfmWithOpaqueFallback;

    // Max characters before a simple table becomes opaque.
    public int TableInlineCellMax = 80;

    // Tracked changes: accept silently, render as {+/-}, or strip dels.
    public TrackedChangeMode TrackedChanges = TrackedChangeMode.Accept;

    // Resolve list numbering to literal markers (default true).
    public bool ResolveNumbering = true;

    // Custom image URI scheme. Default: "docxodus://img/{unid}"
    public Func<ImageInfo, string>? ImageUriBuilder;
}

[Flags]
public enum ProjectionScopes
{
    Body = 1, Headers = 2, Footers = 4, Footnotes = 8, Endnotes = 16, Comments = 32,
    All = Body | Headers | Footers | Footnotes | Endnotes | Comments
}

public enum AnchorRenderMode { Block, BlockAndInline, None }
public enum TableRenderMode { GfmWithOpaqueFallback, AlwaysGfm, AlwaysOpaque }
public enum TrackedChangeMode { Accept, RenderInline, StripDeletions }
```

## Implementation Plan (Phases)

1. **Anchor index.** Walk the document, assign/reuse Unids on every block-level element across all parts, build the `AnchorIndex`. No markdown output yet — just verify uniqueness and round-trip resolution.
2. **Plain paragraphs + headings.** Emit body paragraphs and styled headings with anchors. Tests against `TestFiles/HC*` documents.
3. **Inline runs.** Bold, italic, code, strike, hyperlinks.
4. **Lists with resolved numbering.** Lean on existing `ListItemRetrieverSettings` infrastructure.
5. **Simple tables (GFM) + opaque fallback.**
6. **Multipart parts.** Headers, footers, footnotes, endnotes, comments.
7. **Tracked changes rendering modes.**
8. **WASM + npm wrapper.** Add `[JSExport]` methods and TypeScript types matching the other converters.

Each phase ships with tests in `Docxodus.Tests/WmlToMarkdownConverterTests.cs` (test prefix `MD###`).

## Open Questions

- **Anchor stability across re-serialization.** Today Unids are assigned when needed; we should verify they survive `OpenXmlMemoryStreamDocument` round trips and document the lifecycle. If they don't, the converter must persist them back to the document.
- **Comment-on-span granularity.** Word comments can anchor to a span that crosses runs and paragraphs. Inline `{#cmt:…}` markers handle intra-paragraph; cross-paragraph spans probably need a start/end pair.
- **Images.** The placeholder `docxodus://img/{unid}` URI scheme assumes the caller provides resolution. Should we instead emit data URIs? Configurable via `ImageUriBuilder`, default TBD.
- **Bidirectional editing.** This converter is one-way, but the eventual `MarkdownToWml` story (or, more likely, an `ApplyEdit(anchor, op)` API) needs to be designed alongside its callers — not in isolation here.

## Related

- [`docx_converter.md`](docx_converter.md) — `WmlToHtmlConverter` internals (the sibling converter this most resembles)
- [`opencontracts_export.md`](opencontracts_export.md) — Other "structured export" precedent
- [`incremental_annotation_overlay.md`](incremental_annotation_overlay.md) — Anchor-based overlay pattern used by `ExternalAnnotationProjector`
- [`tracked_changes.md`](tracked_changes.md) — How tracked changes inform the `TrackedChangeMode` setting
