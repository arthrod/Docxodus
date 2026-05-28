# Note-Ref Surgery, List Write Tools, and Block Metadata for `DocxSession`

**Date:** 2026-05-28
**Branches:** One per sub-feature (see Implementation order). Designed on `main` after the in-flight `feat/annotation-write-surface` PR lands.
**Status:** Design — awaiting implementation plan.

## Problem

Three gaps in the `DocxSession` surface block real editing pipelines:

1. **Footnote / endnote references can only be deleted via the body.**
   `DocxSession.DeleteBlock` on a `fn`/`en` anchor already removes the note
   definition AND strips every cross-reference (`w:footnoteReference`,
   `w:endnoteReference`) pointing at it. There is no inverse: no way to
   list every ref by location, surgically delete a single citation while
   leaving the body, or clean up orphan bodies after refs are gone.

2. **List write surface is sparse.** Today `DocxSession` exposes only
   `SetListLevel(anchorId, delta)` and `RemoveListMembership(anchorId)`.
   Agents cannot:
   - create a numbered list from scratch
   - convert existing paragraphs to a numbered list
   - restart numbering at a chosen value
   - jump to an absolute level (only delta)
   - move a paragraph between lists (set numId)

3. **List metadata is invisible.** `AnchorInfo` reports `AutoNumberPrefix`
   but hides the structural facts that drove it — `numId`, `ilvl`,
   `abstractNumId`, format (`decimal`/`bullet`/…), whether numbering is
   inherited from the paragraph style or set inline, what start override
   (if any) applies, what `pStyle` resolved name is. Without these, an
   agent that just read a doc cannot reason about how to mutate it
   without round-tripping through `Raw.GetXml`.

This spec adds (a) anchor-scoped note-ref enumeration and removal, (b) a
list write surface covering create / convert / restart / absolute-level /
set-numId, and (c) a metadata read surface that exposes list
membership, block-level style facts, and the enclosing `sectPr`.

## Goals

- **Anchor-addressed throughout.** Every new method targets existing
  markdown-projection anchors. No new targeting modes; callers reuse
  `Grep` / `FindByText` / `FindByKind` to obtain anchors.
- **One return shape.** `EditResult` envelope for every mutation, with
  one new optional field (`NoteRefSummary`) to carry per-call removal
  verification.
- **Caller-controlled numbering identity.** Creating or converting to a
  numbered list accepts an optional `numId`. Null means create a fresh
  `w:num` over a session-reused canonical abstractNum for the requested
  format. Provided means use that num (validated against numbering.xml,
  format mismatch rejected explicitly).
- **Undo/redo for free.** Each mutation snapshots the package into the
  existing undo ring; one `Undo()` reverses a whole call (including
  bulk-range note-ref removal).
- **Symmetric public surface.** All public methods live flat on
  `DocxSession` (matching every other tier). Implementation factored
  into `Internal/NoteRefOps.cs`, `Internal/ListWriteOps.cs`, and
  `Internal/BlockMetadataOps.cs` — the same factoring used for
  `Internal/AnnotationOps.cs`.

## Non-goals

- **No auto ↔ literal-prefix conversion.** Detecting a literal prefix
  in run text reliably (vs. authored body content that happens to look
  like a label) is fragile. Out of scope; pursue separately if a use
  case emerges.
- **No bulleted-list creation parity beyond `NumberFormat.Bullet`.**
  We support `Bullet` as a format value (caller-supplied bullet glyph
  via the abstractNum); we don't add a separate `InsertBulletedList`.
- **No table-cell metadata.** Row/column index, gridSpan/vMerge,
  shading, borders warrant their own design — explicitly deferred.
- **No tracked-change rendering for list write ops.** OOXML has no
  native track-changes markup for numbering edits. With
  `TrackedChanges = RenderInline`, list write ops still mutate
  immediately (documented, no error).
- **No changes to .NET CLI tools or the web demo.**

## Architecture

Three feature areas, sharing the same cross-layer rollout pattern (core →
ops facade → JSON helpers → WASM bridge / stdio dispatcher → npm wrapper
→ Python wrapper → tests + docs).

### Feature area 1 — Note refs

`Internal/NoteRefOps.cs` is the new home for ref enumeration and removal.
Public methods on `DocxSession`:

```csharp
public IReadOnlyList<NoteReference> ListNoteReferences();
public IReadOnlyList<NoteReference> ListNoteReferences(string anchorId);
public IReadOnlyList<NoteReference> ListNoteReferencesInRange(
    string fromAnchorId, string toAnchorIdExclusive);
public IReadOnlyList<NoteReference> ListNoteReferencesInSection(
    string headingAnchorId);

public EditResult RemoveNoteReferences(
    IEnumerable<string> runUnids, bool deleteOrphanedBodies);
public EditResult RemoveNoteReferencesInAnchor(
    string anchorId, bool deleteOrphanedBodies);
public EditResult RemoveNoteReferencesInRange(
    string fromAnchorId, string toAnchorIdExclusive, bool deleteOrphanedBodies);
public EditResult RemoveNoteReferencesInSection(
    string headingAnchorId, bool deleteOrphanedBodies);
```

`NoteReference` records the noteId, kind, enclosing anchor, character
span, host run unid, note-body preview, and a `BodyExists` flag (false
for orphan refs whose body is missing). Removal identifies refs by host
run unid (stable across mutations) so callers can address a single
citation precisely. The body-side cleanup is opt-in via
`deleteOrphanedBodies` — when the last ref to a body is removed and the
flag is true, the body goes too; otherwise the body is reported in
`NoteRefSummary.OrphanedNoteIds`.

### Feature area 2 — List writes

`Internal/ListWriteOps.cs` covers numbering identity discovery and the
write methods. Public methods on `DocxSession`:

```csharp
public IReadOnlyList<ListDefinition> ListListDefinitions();

public EditResult InsertNumberedList(
    string atAnchorId, Position position, IReadOnlyList<string> itemMarkdowns,
    NumberFormat format, int? startN = null, int? numId = null);

public EditResult ConvertToNumberedList(
    string anchorId, NumberFormat format,
    int? startN = null, int? numId = null);
public EditResult ConvertRangeToNumberedList(
    string fromAnchorId, string toAnchorIdExclusive,
    NumberFormat format, int? startN = null, int? numId = null);

public EditResult RestartNumberedList(string anchorId, int startN);
public EditResult SetListLevelAbsolute(string anchorId, int level);
public EditResult SetListNumId(string anchorId, int numId);
```

`NumberFormat` is a closed enum: `Decimal, UpperLetter, LowerLetter,
UpperRoman, LowerRoman, Bullet`. Fresh-num creation reuses a single
session-managed canonical `w:abstractNum` per format value (created
lazily the first time the format is requested), so repeated
`InsertNumberedList(format: Decimal)` calls don't bloat numbering.xml
with duplicate abstractNums but each call still produces an independent
`w:num` so counters don't collide.

`startN` is realized as a `w:lvlOverride/w:startOverride` on whichever
num is being used. `ConvertToNumberedList` rejects paragraphs that
already carry a different `w:numPr` (callers must `RemoveListMembership`
first — explicit over silent overwrite); converting a paragraph that
already matches the target (same `numId`, level 0) is a no-op success.

### Feature area 3 — Block metadata

`Internal/BlockMetadataOps.cs` is read-only inspection. Public methods:

```csharp
public BlockMetadata? GetBlockMetadata(string anchorId);
public IReadOnlyDictionary<string, BlockMetadata?> GetBlockMetadatas(
    IEnumerable<string> anchorIds);
public ListMembership? GetListMembership(string anchorId);
public SectionInfo? GetSectionInfo(string anchorId);
```

`BlockMetadata` carries `AnchorId`, `Kind`, `Scope`, `StyleId`,
`StyleName`, `OutlineLevel`, an optional `ListMembership` subrecord, and
`HasInlineFormatting` (true if any descendant `w:r/w:rPr` carries
bold/italic/color/etc.).

`ListMembership` carries `NumId`, `AbstractNumId`, `Level`, `Format`,
optional `StartOverride`, `IsAutoNumbered`, `FromStyle` (true when
numPr is inherited from the paragraph style chain rather than inline),
and `GeneratedLabel` (the same string already exposed as
`AnchorInfo.AutoNumberPrefix` — duplicated here so callers don't have
to take two calls).

`SectionInfo` finds the next `w:sectPr` forward from the anchor (or
the body's trailing sectPr) and reports page size in twips, landscape
orientation, margins in twips, column count, and the URIs of the
header / footer parts the section references. Returns null for anchors
outside the body part (fn/en/hdr/ftr/cmt — sectPr is body-only).

## Error handling

New `EditErrorCode` values:

```csharp
NoteRefNotFound,        // run unid doesn't host a w:footnoteReference/w:endnoteReference
UnknownNumId,           // numId provided doesn't exist in numbering.xml
NumIdFormatMismatch,    // caller-supplied numId resolves to a different NumberFormat
```

Existing codes reused as-is: `AnchorNotFound`, `AnchorWrongKind`,
`AnchorsNotAdjacent`, `InvalidListLevel`, `SessionDisposed`.

**Note refs:** Reserved Word boilerplate (`w:type="separator"` /
`continuationSeparator`) is skipped on both list and remove (no-op,
not an error). Orphan refs (id has no body) list with `BodyExists=false`
and remove cleanly. When the host `w:r` contains only the reference
element, the empty run is also stripped (matches existing
`RemoveEmptyRunIfNeeded` behavior). Each Remove* call records ONE undo
snapshot covering all ref + body removals it performs.

**List writes:** `numId=N` validated against numbering.xml on every
call. Format conflict (`Decimal` requested but `numId` resolves to
`LowerLetter`) → `NumIdFormatMismatch` (the caller must drop the
format arg or pass a matching numId — we never silently override the
existing format). `ConvertRangeToNumberedList` silently skips
non-paragraph siblings (tables, sectPr-only paragraphs) rather than
erroring. `SetListLevelAbsolute` rejects out-of-range with
`InvalidListLevel` (same [0,8] window as `SetListLevel`).
`RestartNumberedList` is idempotent — re-running replaces the prior
`w:startOverride` on that level.

**Block metadata:** Pure reads. Unknown anchors return null (bulk
variant maps each id to null, matching `GetAnchorInfos`).

## Cross-layer ripple

Each new method ripples through the full stack in the order CLAUDE.md
specifies (`DocxSessionOps` first, then bridges + clients):

1. **`Internal/DocxSessionOps.cs`** — one static facade method per
   public API, returning JSON strings via `DocxSessionJson`.
2. **`Internal/DocxSessionJson.cs`** — serializers for `NoteReference`,
   `NoteRefSummary`, `ListDefinition`, `BlockMetadata`,
   `ListMembership`, `SectionInfo`; extend `EditResult` serializer to
   emit `noteRefSummary` when non-null.
3. **WASM bridge — `wasm/DocxodusWasm/DocxSessionBridge.cs`** — one
   `[JSExport]` shell per Ops method.
4. **Stdio host — `tools/python-host/Dispatcher.cs`** — one NDJSON
   `case` per method.
5. **npm — `npm/src/types.ts` + `npm/src/index.ts` +
   `docxodus.worker.ts` + `worker-proxy.ts`** — TS types, wrapper
   methods, worker round-trip.
6. **Python — `python/src/docx_scalpel/types.py` +
   `python/src/docx_scalpel/session.py`** — dataclasses and session
   methods that subprocess to docxodus-pyhost.

## Testing

Three new xUnit files in `Docxodus.Tests/`:

- **`DocxSessionNoteRefTests.cs`** (~25 tests): listing per
  anchor / range / section / document; reserved-note skip; orphan-ref
  detection; partial-removal (multi-cited note keeps body); last-ref
  removal with `deleteOrphanedBodies=true` cleans up; empty-run strip;
  range/section atomicity; cross-parent rejection; persistence
  round-trip via Save+reopen; undo/redo.
- **`DocxSessionListWriteTests.cs`** (~30 tests):
  `ListListDefinitions`; fresh-num path creates one new `w:num` per
  call and reuses one `w:abstractNum` per format across calls;
  caller-supplied numId path; `UnknownNumId` and
  `NumIdFormatMismatch`; `startN` becomes `w:startOverride`;
  `InsertNumberedList` Position.Before/After; rejection of
  non-paragraph anchors; `ConvertToNumberedList` (single + range);
  rejection of paragraphs with conflicting numPr; mixed-content range
  (tables skipped); `RestartNumberedList` idempotent re-run;
  `SetListLevelAbsolute` boundary + out-of-range; `SetListNumId`
  rejects bad numId; undo/redo + Save+reopen for each.
- **`DocxSessionMetadataTests.cs`** (~20 tests):
  `GetBlockMetadata` for paragraph / heading / list item / table cell;
  `ListMembership` for inline vs. style-inherited numPr; `FromStyle`
  semantics; `StartOverride` capture; `GetSectionInfo` body anchor
  finds next sectPr; landscape / columns / margins / header-footer
  uris; null for non-body anchors; bulk variant dedup + order.

Plus cross-layer parity (matching the annotation-write rollout):

- **npm Playwright** — one spec per feature area in `npm/tests/`
  exercising the worker proxy round-trip.
- **Python pytest** — coverage per feature area in `python/tests/`
  invoking docxodus-pyhost end-to-end.

Test data: reuse existing `TestFiles/` documents where possible (the
comparer/converter suites already include rich list and footnote
samples). For numbering tests, build minimal programmatic docs
(matching `DocumentBuilderTests` style) when reuse doesn't fit — gives
control over specific `numId` / `abstractNumId` values needed for
assertions.

## Documentation

- **`CHANGELOG.md`** — entries under `[Unreleased]` for each feature
  area.
- **`CLAUDE.md`** — extend the `DocxSession` bullet list in the
  architecture overview with the new tier (Tier F: note references)
  and the new list write ops + metadata reads.
- **`docs/architecture/docx_mutation_api.md`** — three new sections:
  - "Tier F: Note references" — covers
    `ListNoteReferences*` / `RemoveNoteReferences*`, the orphan-body
    model, multi-citation semantics, and the `NoteRefSummary` shape.
  - Extension to "Tier C: formatting / lists" — covers
    `InsertNumberedList`, `ConvertToNumberedList`,
    `ConvertRangeToNumberedList`, `RestartNumberedList`,
    `SetListLevelAbsolute`, `SetListNumId`, the numbering-identity
    model (fresh-num + reused canonical abstractNum per format), and
    the `NumberFormat` enum.
  - "Inspection: block metadata" — covers `GetBlockMetadata`,
    `GetBlockMetadatas`, `GetListMembership`, `GetSectionInfo`, and
    the `BlockMetadata` / `ListMembership` / `SectionInfo` records.

## Implementation order

Three sub-features, sequenced so each builds on its dependency
without cross-feature churn:

1. **Block metadata first.** It's pure-read, low-risk, and produces
   the `ListMembership` type that the list write ops verify their
   results against in tests.
2. **List writes second.** Uses `ListMembership` from feature 1 to
   round-trip mutations in tests.
3. **Note refs last.** Independent of the other two; deliberately
   sequenced last so it can land separately if list/metadata
   reviews stretch.

Each sub-feature is its own PR (matching the annotation-write
pattern of one design → one branch → several focused commits → one
PR).
