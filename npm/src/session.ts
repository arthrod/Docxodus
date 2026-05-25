// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

import type {
  AnchorRef,
  CharSpan,
  DocxodusWasmExports,
  DocxSessionProjection,
  DocxSessionSettings,
  EditResult,
  FormatOp,
  GrepOptions,
  ReplaceOptions,
  TemplatePlaceholder,
  TextMatch,
} from "./types.js";
import { PlaceholderKinds } from "./types.js";

/**
 * Stateful in-memory DOCX editing session keyed by markdown-projection anchor ids.
 * Mirror of the .NET `DocxSession` surface. See
 * `docs/architecture/docx_mutation_api.md` for the surface contract,
 * anchor lifecycle, error catalog, and supported markdown subset.
 *
 * Sessions are not eligible for JS-side garbage collection — call {@link close}
 * (or use a `using` block under TypeScript 5.2+) when done.
 */
export class DocxSession {
  private readonly handle: number;
  private readonly wasm: DocxodusWasmExports["DocxSessionBridge"];

  /** @internal */
  constructor(handle: number, wasm: DocxodusWasmExports["DocxSessionBridge"]) {
    this.handle = handle;
    this.wasm = wasm;
  }

  // ─── View ────────────────────────────────────────────────────────────

  project(): DocxSessionProjection {
    return JSON.parse(this.wasm.Project(this.handle)) as DocxSessionProjection;
  }

  // ─── Tier A: text CRUD ───────────────────────────────────────────────

  replaceText(anchorId: string, markdown: string): EditResult {
    return JSON.parse(this.wasm.ReplaceText(this.handle, anchorId, markdown)) as EditResult;
  }

  deleteBlock(anchorId: string): EditResult {
    return JSON.parse(this.wasm.DeleteBlock(this.handle, anchorId)) as EditResult;
  }

  // ─── Tier B: structural ──────────────────────────────────────────────

  insertParagraph(anchorId: string, position: "before" | "after", markdown: string): EditResult {
    return JSON.parse(this.wasm.InsertParagraph(this.handle, anchorId, position, markdown)) as EditResult;
  }

  splitParagraph(anchorId: string, characterOffset: number): EditResult {
    return JSON.parse(this.wasm.SplitParagraph(this.handle, anchorId, characterOffset)) as EditResult;
  }

  mergeParagraphs(firstAnchorId: string, secondAnchorId: string): EditResult {
    return JSON.parse(this.wasm.MergeParagraphs(this.handle, firstAnchorId, secondAnchorId)) as EditResult;
  }

  // ─── Tier C: formatting ──────────────────────────────────────────────

  applyFormat(anchorId: string, span: CharSpan | null, op: FormatOp): EditResult {
    const spanJson = span ? JSON.stringify(span) : "";
    return JSON.parse(this.wasm.ApplyFormat(this.handle, anchorId, spanJson, JSON.stringify(op))) as EditResult;
  }

  /**
   * Convenience: find `substring` in the anchor's flat text and apply `op` to the
   * first occurrence. Eliminates the offset-arithmetic trap from #138 — caller passes
   * the visible text they want formatted, the WASM-side resolves it to a CharSpan.
   */
  applyFormatBySubstring(anchorId: string, substring: string, op: FormatOp): EditResult {
    return JSON.parse(
      this.wasm.ApplyFormatBySubstring(this.handle, anchorId, substring, JSON.stringify(op))
    ) as EditResult;
  }

  /**
   * Convenience: apply `op` to the exact span of a {@link TextMatch} (typically from
   * {@link grep}). The match's `enclosingAnchor.id` + `span` address one specific
   * occurrence even when several identical needles share the same block.
   */
  applyFormatToMatch(match: TextMatch, op: FormatOp): EditResult {
    const span: CharSpan = { start: match.span.start, length: match.span.length };
    return this.applyFormat(match.enclosingAnchor.id, span, op);
  }

  setParagraphStyle(anchorId: string, styleId: string): EditResult {
    return JSON.parse(this.wasm.SetParagraphStyle(this.handle, anchorId, styleId)) as EditResult;
  }

  setListLevel(anchorId: string, levelDelta: number): EditResult {
    return JSON.parse(this.wasm.SetListLevel(this.handle, anchorId, levelDelta)) as EditResult;
  }

  removeListMembership(anchorId: string): EditResult {
    return JSON.parse(this.wasm.RemoveListMembership(this.handle, anchorId)) as EditResult;
  }

  // ─── Tier D: cell content ────────────────────────────────────────────

  replaceCellContent(cellAnchorId: string, markdown: string): EditResult {
    return JSON.parse(this.wasm.ReplaceCellContent(this.handle, cellAnchorId, markdown)) as EditResult;
  }

  // ─── Raw escape hatch ────────────────────────────────────────────────

  readonly raw = {
    getXml: (anchorId: string): string => this.wasm.RawGetXml(this.handle, anchorId),
    insertXml: (anchorId: string, position: "before" | "after", xml: string): EditResult =>
      JSON.parse(this.wasm.RawInsertXml(this.handle, anchorId, position, xml)) as EditResult,
    replaceXml: (anchorId: string, xml: string): EditResult =>
      JSON.parse(this.wasm.RawReplaceXml(this.handle, anchorId, xml)) as EditResult,
  };

  // ─── Search ──────────────────────────────────────────────────────────

  /**
   * Searches the flat text of every paragraph/heading/list-item in scope for
   * matches of `pattern`, returning them in document order with the run
   * fragments each match spans. Lets callers rewrite a match in place while
   * preserving each fragment's formatting (bold/italic/hyperlink/etc.).
   *
   * `pattern` is a regular expression — use plain string equivalents wrapped
   * in `^` / `$` or pass literal text escaped via a helper.
   *
   * @see docs/architecture/docx_mutation_api.md#grep
   */
  grep(pattern: string, options?: GrepOptions): TextMatch[] {
    return JSON.parse(this.wasm.Grep(this.handle, pattern, options ? JSON.stringify(options) : "")) as TextMatch[];
  }

  /**
   * Finds every literal occurrence of `find` in the anchor's flat text and
   * replaces it with `replace`, preserving the surrounding run formatting that
   * the match didn't touch. Returns one `EditResult` per attempted match.
   *
   * Run-formatting contract: the replacement text inherits the formatting of
   * the FIRST run the match spanned. Middle/trailing runs keep their `w:rPr`
   * but lose the slice of text the match consumed.
   *
   * @see docs/architecture/docx_mutation_api.md#replacetextrange
   */
  replaceTextRange(anchorId: string, find: string, replace: string, options?: ReplaceOptions): EditResult[] {
    return JSON.parse(
      this.wasm.ReplaceTextRange(this.handle, anchorId, find, replace, options ? JSON.stringify(options) : "")
    ) as EditResult[];
  }

  /**
   * Replaces a specific Grep match in place — addresses the exact span by
   * `enclosingAnchor.id` + `span.{start,length}`, so identical needles in the
   * same paragraph (the template-fill case where five `[___]` placeholders
   * each get a different value) don't collide.
   */
  replaceMatch(match: TextMatch, replace: string): EditResult {
    return JSON.parse(
      this.wasm.ReplaceTextAtSpan(this.handle, match.enclosingAnchor.id, match.span.start, match.span.length, replace)
    ) as EditResult;
  }

  /**
   * Enumerate template placeholders in the document. Thin classifier over
   * {@link grep}: distinguishes `[___]` value blanks (`blank_fill`),
   * `[bracketed alternative clauses]` (`alternative_clause`), and
   * `[insert X]` / `[*italic hint*]` instructions (`instruction`).
   *
   * Combine kinds with bitwise OR: `PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction`.
   * Default is `PlaceholderKinds.All`; default scope is body only (1).
   *
   * @see docs/architecture/docx_mutation_api.md#findplaceholders
   */
  findPlaceholders(
    kinds: number = PlaceholderKinds.All,
    scope: number = 1,
  ): TemplatePlaceholder[] {
    return JSON.parse(this.wasm.FindPlaceholders(this.handle, kinds, scope)) as TemplatePlaceholder[];
  }

  // ─── Lifecycle ───────────────────────────────────────────────────────

  undo(): boolean {
    return this.wasm.Undo(this.handle);
  }

  redo(): boolean {
    return this.wasm.Redo(this.handle);
  }

  save(): Uint8Array {
    return this.wasm.Save(this.handle);
  }

  close(): void {
    this.wasm.CloseSession(this.handle);
  }

  // TypeScript 5.2+ disposable protocol
  [Symbol.dispose]?(): void {
    this.close();
  }
}

/**
 * Opens a new {@link DocxSession} over the supplied DOCX bytes.
 * The returned session holds its document in WASM memory until you call
 * {@link DocxSession.close} (or it is disposed).
 */
export function openDocxSession(
  bytes: Uint8Array,
  wasmExports: DocxodusWasmExports,
  settings?: DocxSessionSettings,
): DocxSession {
  const bridge = wasmExports.DocxSessionBridge;
  const handle = bridge.OpenSession(bytes, settings ? JSON.stringify(settings) : "");
  return new DocxSession(handle, bridge);
}

export type { AnchorRef, CharSpan, DocxSessionProjection, DocxSessionSettings, EditError, EditErrorCode, EditResult, FormatOp, GrepOptions, MarkdownPatch, PlaceholderKind, ReplaceOptions, RunFormatting, RunFragment, TemplatePlaceholder, TextMatch } from "./types.js";
export { PlaceholderKinds } from "./types.js";
