#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

// ─── Public value types ────────────────────────────────────────────────────

public enum Position { Before, After }

public readonly record struct CharSpan(int Start, int Length);

public sealed record FormatOp
{
    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    public bool? Strike { get; init; }
    public bool? Code { get; init; }
    public string? Color { get; init; }
    public string? RunStyle { get; init; }
}

/// <summary>
/// Per-fragment visible formatting reported by <see cref="DocxSession.Grep"/>.
/// Booleans default to <c>false</c> meaning "not set on this fragment". The
/// fields cover what a callerlikely wants to preserve when rewriting a match in
/// place — character emphasis, color, hyperlink target, named run style.
/// </summary>
public sealed record RunFormatting
{
    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    public bool Code { get; init; }
    public string? Color { get; init; }
    public string? HyperlinkUrl { get; init; }
    public string? RunStyle { get; init; }
}

/// <summary>
/// One piece of a <see cref="TextMatch"/> that came from a single <c>&lt;w:r&gt;</c> run.
/// The <see cref="Unid"/> uniquely identifies the run within the document; callers
/// rewriting the match can address each piece by its Unid + <see cref="SpanInElement"/>
/// and preserve the run's <see cref="Formatting"/> when constructing replacement XML.
/// </summary>
public sealed record RunFragment
{
    /// <summary>PtOpenXml.Unid of the <c>w:r</c> element this fragment came from.</summary>
    required public string Unid { get; init; }

    /// <summary>The text from this run that participates in the match.</summary>
    required public string Text { get; init; }

    /// <summary>Character offset + length of this fragment inside the run's flat text.</summary>
    required public CharSpan SpanInElement { get; init; }

    /// <summary>Visible formatting of the run this fragment came from.</summary>
    required public RunFormatting Formatting { get; init; }
}

/// <summary>
/// A single match returned by <see cref="DocxSession.Grep"/>. The match always lives
/// within one block-level element (the <see cref="EnclosingAnchor"/>); cross-block
/// matches aren't represented because OOXML doesn't allow text to span paragraphs.
/// </summary>
public sealed record TextMatch
{
    /// <summary>The matched text.</summary>
    required public string Text { get; init; }

    /// <summary>The smallest block-level anchor (paragraph/heading/list item/table cell) that fully contains the match.</summary>
    required public AnchorTarget EnclosingAnchor { get; init; }

    /// <summary>Character offset + length of the match in the enclosing block's flat text.</summary>
    required public CharSpan Span { get; init; }

    /// <summary>The run fragments the match spans, in document order. Always non-empty for a successful match.</summary>
    required public IReadOnlyList<RunFragment> Fragments { get; init; }

    /// <summary>Up to <c>contextChars</c> chars from the enclosing block immediately before the match.</summary>
    required public string ContextBefore { get; init; }

    /// <summary>Up to <c>contextChars</c> chars from the enclosing block immediately after the match.</summary>
    required public string ContextAfter { get; init; }

    /// <summary>Regex capture groups (index 0 is always the whole match; named groups appear at their numeric index).</summary>
    public IReadOnlyList<string> Groups { get; init; } = Array.Empty<string>();
}

/// <summary>Options that tune <see cref="DocxSession.ReplaceTextRange"/>.</summary>
public sealed record ReplaceOptions
{
    /// <summary>Case-insensitive matching for the literal <c>find</c> needle.</summary>
    public bool IgnoreCase { get; init; }

    /// <summary>Cap the number of replacements; null = unlimited.</summary>
    public int? MaxReplacements { get; init; }
}

public sealed record AnchorInfo(string Id, string Kind, string Scope, string TextPreview);

public sealed record MarkdownPatch(string ScopeAnchorId, string Markdown);

public sealed record EditError(EditErrorCode Code, string Message, string? AnchorId = null);

public enum EditErrorCode
{
    AnchorNotFound,
    AnchorWrongKind,
    AnchorsNotAdjacent,
    SessionDisposed,

    MalformedMarkdown,
    UnsupportedMarkdownSyntax,
    TableInsertNotSupported,
    FootnoteRefNotSupported,
    CommentMarkerNotSupported,
    ImageInsertNotSupported,
    AnchorTokenInPayload,

    OffsetOutOfRange,
    InvalidPosition,

    UnknownStyle,
    InvalidListLevel,

    MalformedXml,
    DisallowedNamespace,
    IncompatibleElementType,
    ValidationFailed,

    NothingToUndo,
    NothingToRedo,

    InternalError,
}

public sealed class EditResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public IReadOnlyList<Anchor> Created { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Removed { get; init; } = Array.Empty<Anchor>();
    public IReadOnlyList<Anchor> Modified { get; init; } = Array.Empty<Anchor>();
    public MarkdownPatch? Patch { get; init; }

    internal static EditResult Fail(EditErrorCode code, string message, string? anchorId = null) =>
        new() { Success = false, Error = new EditError(code, message, anchorId) };
}

public sealed class DocxSessionSettings
{
    public int UndoDepth { get; init; } = 50;
    public bool ValidateRawOps { get; init; } = false;
    public TrackedChangeMode TrackedChanges { get; init; } = TrackedChangeMode.Accept;
    public string? RevisionAuthor { get; init; }
    public WmlToMarkdownConverterSettings ProjectionSettings { get; init; } = new();

    /// <summary>
    /// When <c>false</c> (default) <see cref="DocxSession.Save"/> strips
    /// <c>PtOpenXml:Unid</c> attributes from every part — the attribute is internal
    /// to the projector and not in the OOXML schema, so persisting it bloats saved
    /// DOCX files (a 100-page document grows by ~700 KB of attribute noise). Set to
    /// <c>true</c> when anchor ids must survive a save/reopen round trip — the
    /// scenario flagged by Open Question #1 in <c>docs/architecture/markdown_projection.md</c>.
    /// </summary>
    public bool PersistAnchorIds { get; init; } = false;
}

// ─── Session ───────────────────────────────────────────────────────────────

public sealed class DocxSession : IDisposable
{
    private readonly DocxSessionSettings _settings;
    private readonly Internal.UndoRing<DocumentSnapshot> _history;
    private MemoryStream? _stream;
    private WordprocessingDocument? _doc;
    private MarkdownProjection? _cachedProjection;
    private bool _disposed;
    private int _revisionCounter = 1000;
    private RawDocxOps? _raw;

    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        _settings = settings ?? new DocxSessionSettings();
        _history = new Internal.UndoRing<DocumentSnapshot>(_settings.UndoDepth);
        _stream = new MemoryStream();
        _stream.Write(docxBytes, 0, docxBytes.Length);
        _stream.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);
    }

    public Exception? LastInternalError { get; private set; }

    public MarkdownProjection Project()
    {
        ThrowIfDisposed();
        return _cachedProjection ??=
            WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
    }

    /// <summary>
    /// Looks up an anchor id with a fallback to Unid-only resolution. The dictionary
    /// is keyed by full <c>kind:scope:unid</c> id, so when a mutation flips the kind
    /// prefix (e.g., <c>p:body:abcd</c> → <c>h:body:abcd</c> after promoting to a
    /// heading), a cached old id would otherwise miss. This helper trails through
    /// to a Unid scan, so agents that hold cached ids keep working — matching the
    /// promise in <c>docs/architecture/docx_mutation_api.md</c>.
    /// </summary>
    internal AnchorTarget? FindAnchor(string? anchorId)
    {
        if (anchorId is null) return null;
        var index = Project().AnchorIndex;
        if (index.TryGetValue(anchorId, out var direct)) return direct;
        int lastColon = anchorId.LastIndexOf(':');
        if (lastColon <= 0 || lastColon == anchorId.Length - 1) return null;
        var unid = anchorId.Substring(lastColon + 1);
        foreach (var v in index.Values)
        {
            if (v.Unid == unid) return v;
        }
        return null;
    }

    public bool Exists(string anchorId)
    {
        ThrowIfDisposed();
        return FindAnchor(anchorId) is not null;
    }

    public AnchorInfo? GetAnchorInfo(string anchorId)
    {
        ThrowIfDisposed();
        var target = FindAnchor(anchorId);
        if (target is null) return null;

        var element = target.Resolve(_doc!);
        var preview = element is null ? "" : ElementTextPreview(element);
        return new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, preview);
    }

    /// <summary>
    /// Searches the flat text of every paragraph/heading/list-item in <paramref name="scope"/>
    /// for matches of <paramref name="pattern"/> and returns them in document order, each
    /// with the run fragments it spans. The fragment list lets callers rewrite a match in
    /// place while preserving each fragment's formatting — see #143 for design context.
    /// </summary>
    /// <param name="pattern">Regular-expression pattern (use <c>Regex.Escape</c> for literal text).</param>
    /// <param name="options">Standard <see cref="System.Text.RegularExpressions.RegexOptions"/> flags.</param>
    /// <param name="scope">Which package parts to search. Defaults to <see cref="ProjectionScopes.Body"/>.</param>
    /// <param name="contextChars">Number of characters of surrounding text to include in
    /// <see cref="TextMatch.ContextBefore"/> and <see cref="TextMatch.ContextAfter"/>.</param>
    public IReadOnlyList<TextMatch> Grep(
        string pattern,
        System.Text.RegularExpressions.RegexOptions options = System.Text.RegularExpressions.RegexOptions.None,
        ProjectionScopes scope = ProjectionScopes.Body,
        int contextChars = 40)
    {
        ThrowIfDisposed();
        if (string.IsNullOrEmpty(pattern)) return Array.Empty<TextMatch>();

        var regex = new System.Text.RegularExpressions.Regex(pattern, options);
        var results = new List<TextMatch>();

        // Walk the projection's AnchorIndex so document order is the same order
        // an agent sees in the projection. Only block-level kinds that hold runs
        // qualify (paragraphs/headings/list-items/table cells); other kinds either
        // don't contain text directly (tbl, tr, sec) or live in non-body scopes
        // we filter explicitly below.
        var index = Project().AnchorIndex;
        foreach (var target in index.Values)
        {
            if (!ScopeMatches(target.Anchor.Scope, scope)) continue;
            if (target.Anchor.Kind is not ("p" or "h" or "li" or "tc")) continue;

            var element = target.Resolve(_doc!);
            if (element is null) continue;

            // Table cells contain paragraphs; recurse so a Grep over the body
            // also hits cell text. Other kinds operate on the element directly.
            if (target.Anchor.Kind == "tc")
            {
                // Cell paragraphs are reachable via their own AnchorIndex entries,
                // so skip the cell wrapper to avoid double-counting matches.
                continue;
            }

            var map = Internal.RunTextMap.Build(element);
            if (map.FlatText.Length == 0) continue;

            // Look up the owner part once per anchor so the hyperlink resolver
            // doesn't have to walk back up to the root annotation per run.
            var ownerPart = ResolvePart(target.PartUri);

            foreach (System.Text.RegularExpressions.Match m in regex.Matches(map.FlatText))
            {
                if (!m.Success || m.Length == 0) continue;

                var pieces = Internal.RunTextMap.ResolveRange(map, m.Index, m.Length);
                if (pieces.Count == 0) continue;

                var fragments = new List<RunFragment>(pieces.Count);
                foreach (var (seg, offsetInRun, len) in pieces)
                {
                    var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
                    var runText = RunText(seg.Run);
                    fragments.Add(new RunFragment
                    {
                        Unid = runUnid,
                        Text = runText.Substring(offsetInRun, len),
                        SpanInElement = new CharSpan(offsetInRun, len),
                        Formatting = ExtractFormatting(seg.Run, ownerPart),
                    });
                }

                var ctxStart = Math.Max(0, m.Index - contextChars);
                var ctxBefore = map.FlatText.Substring(ctxStart, m.Index - ctxStart);
                var ctxEnd = Math.Min(map.FlatText.Length, m.Index + m.Length + contextChars);
                var ctxAfter = map.FlatText.Substring(m.Index + m.Length, ctxEnd - (m.Index + m.Length));

                var groups = new string[m.Groups.Count];
                for (int i = 0; i < m.Groups.Count; i++) groups[i] = m.Groups[i].Value;

                results.Add(new TextMatch
                {
                    Text = m.Value,
                    EnclosingAnchor = target,
                    Span = new CharSpan(m.Index, m.Length),
                    Fragments = fragments,
                    ContextBefore = ctxBefore,
                    ContextAfter = ctxAfter,
                    Groups = groups,
                });
            }
        }

        return results;
    }

    /// <summary>
    /// Surgical text replacement within a single paragraph/heading/list-item: finds every
    /// literal occurrence of <paramref name="find"/> in the anchor's flat text and replaces
    /// it with <paramref name="replace"/>, preserving the surrounding run formatting that
    /// the match didn't touch. Returns one <see cref="EditResult"/> per attempted match.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The replacement text is plain-text and inherits the formatting of the FIRST run the
    /// match spanned — middle/trailing runs keep their <c>w:rPr</c> but lose the slice of
    /// text the match consumed (so a bold run that contributed three chars to the match now
    /// has those three chars gone, but stays bold for everything else it held).
    /// </para>
    /// <para>
    /// Matches are applied in reverse document order so multiple matches in the same
    /// paragraph don't invalidate each other's offsets. The whole call records a single undo
    /// snapshot — <see cref="Undo"/> rolls back every replacement together.
    /// </para>
    /// </remarks>
    public IReadOnlyList<EditResult> ReplaceTextRange(
        string anchorId,
        string find,
        string replace,
        ReplaceOptions? options = null)
    {
        if (_disposed)
            return new[] { EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed") };
        if (string.IsNullOrEmpty(find))
            return new[] { EditResult.Fail(EditErrorCode.MalformedMarkdown, "find must be non-empty", anchorId) };

        var target = FindAnchor(anchorId);
        if (target is null)
            return new[] { EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId) };
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return new[] { EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceTextRange requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId) };

        var opts = options ?? new ReplaceOptions();
        var regexOpts = opts.IgnoreCase
            ? System.Text.RegularExpressions.RegexOptions.IgnoreCase
            : System.Text.RegularExpressions.RegexOptions.None;
        var pattern = System.Text.RegularExpressions.Regex.Escape(find);

        var matches = Grep(pattern, regexOpts)
            .Where(m => m.EnclosingAnchor.Anchor.Id == target.Anchor.Id)
            .ToList();
        if (opts.MaxReplacements is int cap) matches = matches.Take(cap).ToList();
        if (matches.Count == 0) return Array.Empty<EditResult>();

        var element = target.Resolve(_doc!);
        if (element is null)
            return new[] { EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId) };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Reverse offset order so earlier-offset matches' SpanInElement stays valid
            // after later-offset edits land — see DS112/DS115.
            foreach (var match in matches.OrderByDescending(m => m.Span.Start))
                ApplyFragmentReplacement(element, match, replace);

            InvalidateProjectionCache();
            var success = new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
            return Enumerable.Repeat(success, matches.Count).ToArray();
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return new[] { EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId) };
        }
    }

    /// <summary>
    /// Convenience: replace a single <see cref="TextMatch"/> (typically from <see cref="Grep"/>)
    /// in place with <paramref name="replace"/>. Same fragment-formatting semantics as
    /// <see cref="ReplaceTextRange"/>.
    /// </summary>
    public EditResult ReplaceMatch(TextMatch match, string replace)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (match is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "match is null");
        return ReplaceTextAtSpan(match.EnclosingAnchor.Anchor.Id, match.Span.Start, match.Span.Length, replace);
    }

    /// <summary>
    /// Surgical replacement of an exact byte range within one block's flat text.
    /// The natural pair to <see cref="Grep"/>: pass the <see cref="TextMatch.EnclosingAnchor"/>'s
    /// id plus the <see cref="TextMatch.Span"/> coordinates to replace one specific match
    /// even when several identical needles share the same paragraph (the template-filling
    /// case where five <c>[___]</c> placeholders each get a different value).
    /// </summary>
    public EditResult ReplaceTextAtSpan(string anchorId, int spanStart, int spanLength, string replace)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceTextAtSpan requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var map = Internal.RunTextMap.Build(element);
        if (spanStart < 0 || spanLength < 0 || spanStart + spanLength > map.FlatText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span {spanStart}+{spanLength} out of [0, {map.FlatText.Length}]", anchorId);

        var pieces = Internal.RunTextMap.ResolveRange(map, spanStart, spanLength);
        if (pieces.Count == 0)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange, "span resolved to no runs", anchorId);

        // Synthesize fragments from the resolved pieces. The replacement helper only
        // reads Unid + SpanInElement, so the other fields are placeholders.
        var fragments = new List<RunFragment>(pieces.Count);
        foreach (var (seg, offsetInRun, len) in pieces)
        {
            var runUnid = (string?)seg.Run.Attribute(PtOpenXml.Unid) ?? string.Empty;
            fragments.Add(new RunFragment
            {
                Unid = runUnid,
                Text = string.Empty,
                SpanInElement = new CharSpan(offsetInRun, len),
                Formatting = new RunFormatting(),
            });
        }
        var synthetic = new TextMatch
        {
            Text = map.FlatText.Substring(spanStart, spanLength),
            EnclosingAnchor = target,
            Span = new CharSpan(spanStart, spanLength),
            Fragments = fragments,
            ContextBefore = string.Empty,
            ContextAfter = string.Empty,
        };

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            ApplyFragmentReplacement(element, synthetic, replace);
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    /// <summary>
    /// Apply <paramref name="match"/>'s fragment list to the live element, inserting
    /// <paramref name="replace"/> into the first fragment's run and removing each
    /// subsequent fragment's slice from its run (preserving each run's rPr).
    /// </summary>
    private static void ApplyFragmentReplacement(XElement blockElement, TextMatch match, string replace)
    {
        if (match.Fragments.Count == 0) return;

        // Build a unid → XElement run lookup once. The run XElements are the live
        // descendants of `blockElement` (walking hyperlink/sdt containers too).
        var runsByUnid = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var run in InlineRuns(blockElement))
        {
            var unid = (string?)run.Attribute(PtOpenXml.Unid);
            if (unid is not null) runsByUnid[unid] = run;
        }

        for (int i = 0; i < match.Fragments.Count; i++)
        {
            var fragment = match.Fragments[i];
            if (!runsByUnid.TryGetValue(fragment.Unid, out var run)) continue;

            var concat = RunText(run);
            var start = fragment.SpanInElement.Start;
            var len = fragment.SpanInElement.Length;
            if (start < 0 || start + len > concat.Length) continue;

            var before = concat.Substring(0, start);
            var after = concat.Substring(start + len);
            var newText = i == 0 ? before + replace + after : before + after;

            // Collapse all w:t descendants in this run into a single w:t with the new text.
            // Loses any inline <w:tab/>/<w:br/> inside the run's text section — they're rare
            // for placeholder slots and supporting them here would balloon the impl. Run's
            // rPr/proofErr siblings are untouched, which is the formatting-preservation contract.
            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                newText));
        }
    }

    private static bool ScopeMatches(string anchorScope, ProjectionScopes filter)
    {
        // Anchor scopes are strings ("body", "hdr1", "ftr2", "fn", "en", "cmt").
        // ProjectionScopes is a flags enum over the same categories.
        if (anchorScope == "body") return filter.HasFlag(ProjectionScopes.Body);
        if (anchorScope.StartsWith("hdr", StringComparison.Ordinal)) return filter.HasFlag(ProjectionScopes.Headers);
        if (anchorScope.StartsWith("ftr", StringComparison.Ordinal)) return filter.HasFlag(ProjectionScopes.Footers);
        if (anchorScope == "fn") return filter.HasFlag(ProjectionScopes.Footnotes);
        if (anchorScope == "en") return filter.HasFlag(ProjectionScopes.Endnotes);
        if (anchorScope == "cmt") return filter.HasFlag(ProjectionScopes.Comments);
        return false;
    }

    private OpenXmlPart? ResolvePart(string partUri) =>
        EnumerateProjectedParts().FirstOrDefault(p => p.Uri.ToString() == partUri);

    private static RunFormatting ExtractFormatting(XElement run, OpenXmlPart? ownerPart)
    {
        var rPr = run.Element(W.rPr);
        string? hyperlinkUrl = null;
        for (var p = run.Parent; p is not null; p = p.Parent)
        {
            if (p.Name == W.hyperlink)
            {
                var rid = (string?)p.Attribute(R.id);
                if (!string.IsNullOrEmpty(rid) && ownerPart is not null)
                {
                    var rel = ownerPart.HyperlinkRelationships.FirstOrDefault(x => x.Id == rid);
                    if (rel is not null) hyperlinkUrl = rel.Uri.ToString();
                }
                break;
            }
        }

        return new RunFormatting
        {
            Bold = rPr?.Element(W.b) is not null,
            Italic = rPr?.Element(W.i) is not null,
            Underline = rPr?.Element(W.u) is not null,
            Strike = rPr?.Element(W.strike) is not null,
            Code = string.Equals((string?)rPr?.Element(W.rStyle)?.Attribute(W.val), "Code", StringComparison.Ordinal),
            Color = (string?)rPr?.Element(W.color)?.Attribute(W.val),
            HyperlinkUrl = hyperlinkUrl,
            RunStyle = (string?)rPr?.Element(W.rStyle)?.Attribute(W.val),
        };
    }

    public byte[] Save()
    {
        ThrowIfDisposed();

        if (_settings.PersistAnchorIds)
        {
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return _stream.ToArray();
        }

        // Strip the internal PtOpenXml:Unid attributes before serializing — they're
        // projector bookkeeping, not OOXML schema, and on a real document the bloat
        // is significant (each Unid is ~50 bytes and the projector assigns one to
        // every descendant of every projected scope). We snapshot first so the
        // session's in-memory state can keep using Unids after the save completes;
        // Project() / Resolve() rely on them.
        var snapshot = TakeSnapshot();
        try
        {
            foreach (var part in EnumerateProjectedParts())
            {
                var xdoc = part.GetXDocument();
                if (xdoc.Root is null) continue;
                bool any = false;
                foreach (var el in xdoc.Root.DescendantsAndSelf())
                {
                    var attr = el.Attribute(PtOpenXml.Unid);
                    if (attr is not null) { attr.Remove(); any = true; }
                }
                if (any) part.PutXDocument();
            }
            _doc!.Save();
            _stream!.Flush();
            _stream.Position = 0;
            return _stream.ToArray();
        }
        finally
        {
            RestoreSnapshot(snapshot);
        }
    }

    /// <summary>
    /// Enumerates every OOXML part the projector walks. Kept centralized so
    /// <see cref="Save"/> (Unid stripping) and any future part-level pass don't drift.
    /// </summary>
    private IEnumerable<OpenXmlPart> EnumerateProjectedParts()
    {
        var main = _doc!.MainDocumentPart;
        if (main is null) yield break;
        yield return main;
        foreach (var h in main.HeaderParts) yield return h;
        foreach (var f in main.FooterParts) yield return f;
        if (main.FootnotesPart is not null) yield return main.FootnotesPart;
        if (main.EndnotesPart is not null) yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null) yield return main.WordprocessingCommentsPart;
    }

    // ─── Tier A: text CRUD ────────────────────────────────────────────────

    public EditResult ReplaceText(string anchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceText requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        // Strip a leading auto-number prefix from the payload before parsing. The
        // projector emits "## Fourth The total number…" — auto-number from numPr
        // plus a space separator plus the run text — so an agent that echoes the
        // visible heading back as its replacement payload otherwise gets the
        // prefix applied twice (Word renders the auto-number AND the run text now
        // begins with "Fourth"). See DS091.
        markdownPayload = StripResolvedAutoNumberPrefix(element, markdownPayload);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_settings.TrackedChanges == TrackedChangeMode.RenderInline)
            {
                ApplyReplaceTextTracked(element, parsed.Blocks);
            }
            else
            {
                ApplyReplaceTextAccept(element, parsed.Blocks);
            }
            PromoteHyperlinkRelationships(element);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult DeleteBlock(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteBlock requires a block-level anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            if (_settings.TrackedChanges == TrackedChangeMode.RenderInline)
            {
                WrapRunsInDel(element);
                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = new[] { target.Anchor },
                    Patch = ProjectScope(target),
                };
            }

            // Collect descendant anchors before removal so the caller knows what's gone.
            var index = Project().AnchorIndex;
            var removed = new List<Anchor> { target.Anchor };
            foreach (var d in element.Descendants())
            {
                var unid = (string?)d.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                foreach (var kv in index)
                {
                    if (kv.Value.Unid == unid && kv.Value.Unid != target.Unid)
                        removed.Add(kv.Value.Anchor);
                }
            }
            element.Remove();
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Removed = removed,
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    // ─── Tier B: structural ops ──────────────────────────────────────────

    public EditResult InsertParagraph(string anchorId, Position pos, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);
        if (parsed.Blocks.Count == 0)
            return EditResult.Fail(EditErrorCode.MalformedMarkdown, "empty payload", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var created = new List<Anchor>();
            var newElements = new List<XElement>();
            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                // List items: try to inherit numbering from a sibling list item so the
                // payload actually projects as a bullet/numbered item. If no sibling
                // has numbering, the paragraph stays bare and the projector classifies
                // it as a plain "p" — which is what we report below.
                if (block.Kind is Internal.ParserBlockKind.BulletItem
                                or Internal.ParserBlockKind.OrderedItem)
                    TryInheritNumPrFromSibling(p, element);
                UnidHelper.AssignToSelfAndDescendants(p);
                newElements.Add(p);
                var unid = (string)p.Attribute(PtOpenXml.Unid)!;
                var kind = ClassifyParagraphKind(p);
                created.Add(new Anchor($"{kind}:{target.Anchor.Scope}:{unid}", kind, target.Anchor.Scope, unid));
            }

            if (pos == Position.Before)
            {
                foreach (var n in newElements) element.AddBeforeSelf(n);
            }
            else
            {
                XElement after = element;
                foreach (var n in newElements) { after.AddAfterSelf(n); after = n; }
            }

            foreach (var n in newElements) PromoteHyperlinkRelationships(n);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult SplitParagraph(string anchorId, int characterOffset)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SplitParagraph requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var totalText = ParagraphText(element);
        if (characterOffset < 0 || characterOffset > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"offset {characterOffset} out of [0, {totalText.Length}]", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            var second = new XElement(W.p);
            if (pPr is not null) second.Add(new XElement(pPr));

            // Split any run that straddles the offset (descends into hyperlinks/sdts),
            // then split any container (hyperlink) that still straddles, then move all
            // inline children + markers at-or-past the offset to `second`.
            SplitRunsAtOffset(element, characterOffset);
            SplitInlineContainersAtOffset(element, characterOffset);
            MoveInlineChildrenAfter(element, characterOffset, second);

            UnidHelper.AssignToSelfAndDescendants(second);
            element.AddAfterSelf(second);

            var secondUnid = (string)second.Attribute(PtOpenXml.Unid)!;
            var secondAnchor = new Anchor(
                $"{target.Anchor.Kind}:{target.Anchor.Scope}:{secondUnid}",
                target.Anchor.Kind, target.Anchor.Scope, secondUnid);

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Created = new[] { secondAnchor },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult MergeParagraphs(string firstAnchorId, string secondAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var firstTarget = FindAnchor(firstAnchorId);
        if (firstTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "first anchor not found", firstAnchorId);
        var secondTarget = FindAnchor(secondAnchorId);
        if (secondTarget is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "second anchor not found", secondAnchorId);

        var firstEl = firstTarget.Resolve(_doc!);
        var secondEl = secondTarget.Resolve(_doc!);
        if (firstEl is null || secondEl is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null");

        if (!ReferenceEquals(firstEl.NextNode, secondEl))
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "MergeParagraphs requires second anchor to be the immediate next sibling of first");

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            // Insert a single-space separator if both sides end/start with non-whitespace.
            // Sentences from two paragraphs should not jam into one another.
            var firstTail = ParagraphText(firstEl);
            var secondHead = ParagraphText(secondEl);
            if (firstTail.Length > 0 && secondHead.Length > 0
                && !char.IsWhiteSpace(firstTail[^1])
                && !char.IsWhiteSpace(secondHead[0]))
            {
                firstEl.Add(new XElement(W.r,
                    new XElement(W.t,
                        new XAttribute(XNamespace.Xml + "space", "preserve"), " ")));
            }

            // Move every paragraph-level child from secondEl into firstEl in document
            // order — runs, hyperlinks, sdts, fldSimples, bookmarkStart/End, comment
            // range markers, etc. The old implementation only moved direct <w:r>
            // children which silently discarded everything else.
            foreach (var child in secondEl.Elements().ToList())
            {
                if (child.Name == W.pPr) continue; // second's pPr is dropped; first's wins
                child.Remove();
                firstEl.Add(child);
            }
            secondEl.Remove();
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { firstTarget.Anchor },
                Removed = new[] { secondTarget.Anchor },
                Patch = ProjectScope(firstTarget),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }

    // ─── Raw escape hatch ────────────────────────────────────────────────

    public RawDocxOps Raw => _raw ??= new RawDocxOps(this);

    private static readonly HashSet<string> AllowedXmlNamespaces = new()
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",        // w:
        "http://schemas.openxmlformats.org/officeDocument/2006/math",          // m:
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", // wp:
        "http://schemas.openxmlformats.org/drawingml/2006/main",               // a:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", // r:
        "http://powertools.codeplex.com/2011",                                 // PtOpenXml (Unid)
    };

    internal string RawGetXmlInternal(string anchorId)
    {
        ThrowIfDisposed();
        var target = FindAnchor(anchorId);
        if (target is null)
            throw new ArgumentException($"anchor not found: {anchorId}");
        var element = target.Resolve(_doc!);
        return element?.ToString() ?? "";
    }

    internal EditResult RawInsertXmlInternal(string anchorId, Position pos, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var (parsedXml, err) = ParseRawXml(xml);
        if (parsedXml is null)
            return new EditResult { Success = false, Error = err! with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        int baselineErrors = _settings.ValidateRawOps ? CountRealValidationErrors() : 0;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            UnidHelper.AssignToSelfAndDescendants(parsedXml);
            if (pos == Position.Before) element.AddBeforeSelf(parsedXml);
            else element.AddAfterSelf(parsedXml);

            if (_settings.ValidateRawOps && CountRealValidationErrors() > baselineErrors)
            {
                var preOp = _history.PopForUndo();
                if (preOp.ok) RestoreSnapshot(preOp.snapshot);
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            InvalidateProjectionCache();
            var freshIndex = Project().AnchorIndex;
            var created = new List<Anchor>();
            foreach (var unid in CollectUnids(parsedXml))
            {
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == unid);
                if (hit is not null) created.Add(hit.Anchor);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    internal EditResult RawReplaceXmlInternal(string anchorId, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var (parsedXml, err) = ParseRawXml(xml);
        if (parsedXml is null)
            return new EditResult { Success = false, Error = err! with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        int baselineErrors = _settings.ValidateRawOps ? CountRealValidationErrors() : 0;
        _history.RecordPreOp(TakeSnapshot());
        try
        {
            UnidHelper.AssignToSelfAndDescendants(parsedXml);
            element.ReplaceWith(parsedXml);

            if (_settings.ValidateRawOps && CountRealValidationErrors() > baselineErrors)
            {
                var preOp = _history.PopForUndo();
                if (preOp.ok) RestoreSnapshot(preOp.snapshot);
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            InvalidateProjectionCache();
            var freshIndex = Project().AnchorIndex;
            var newUnids = CollectUnids(parsedXml).ToHashSet();

            // Classify by Unid set membership: the documented Get→mutate→Replace
            // recipe preserves Unids, so the target anchor must surface as
            // Modified (not as a phantom Removed-then-Created pair). When the
            // replacement XML has fresh Unids — because the caller authored it
            // from scratch — the target is genuinely Removed and the new
            // element(s) are Created. See DS092 / DS092b.
            var modified = new List<Anchor>();
            var removed = new List<Anchor>();
            var created = new List<Anchor>();

            if (newUnids.Contains(target.Unid))
            {
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid);
                if (hit is not null) modified.Add(hit.Anchor);
            }
            else
            {
                removed.Add(target.Anchor);
            }
            foreach (var unid in newUnids)
            {
                if (unid == target.Unid) continue;
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == unid);
                if (hit is not null) created.Add(hit.Anchor);
            }

            return new EditResult
            {
                Success = true,
                Removed = removed,
                Created = created,
                Modified = modified,
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            var preOp = _history.PopForUndo();
            if (preOp.ok) RestoreSnapshot(preOp.snapshot);
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private static (XElement? parsed, EditError? err) ParseRawXml(string xml)
    {
        try
        {
            var x = XElement.Parse(xml);
            foreach (var el in x.DescendantsAndSelf())
            {
                var ns = el.Name.NamespaceName;
                if (!string.IsNullOrEmpty(ns) && !AllowedXmlNamespaces.Contains(ns))
                    return (null, new EditError(EditErrorCode.DisallowedNamespace,
                        $"disallowed namespace: {ns}"));
            }
            return (x, null);
        }
        catch (System.Xml.XmlException ex)
        {
            return (null, new EditError(EditErrorCode.MalformedXml, ex.Message));
        }
    }

    private static IEnumerable<string> CollectUnids(XElement root)
    {
        foreach (var el in root.DescendantsAndSelf())
        {
            var unid = (string?)el.Attribute(PtOpenXml.Unid);
            if (unid is not null) yield return unid;
        }
    }

    // PtOpenXml:Unid is an internal-only attribute added by the projector for anchor
    // addressing; it is not in the OOXML schema, so the validator will emit
    // Sch_UndeclaredAttribute for every occurrence. Filter those out before counting.
    //
    // Mutations operate directly on the part's in-memory XDocument; the validator
    // reads the typed OOXML object model, which is hydrated from the part stream.
    // Flush the XDocument back to the stream first so the validator sees the
    // current state instead of the original document.
    private int CountRealValidationErrors()
    {
        _doc!.MainDocumentPart!.PutXDocument();
        var v = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        return v.Validate(_doc!)
            .Count(e => !(e.Description ?? string.Empty)
                .Contains("http://powertools.codeplex.com/2011", StringComparison.Ordinal));
    }

    // ─── Tier D: table cell content ──────────────────────────────────────

    public EditResult ReplaceCellContent(string cellAnchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(cellAnchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", cellAnchorId);
        if (target.Anchor.Kind != "tc")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ReplaceCellContent requires a cell anchor", cellAnchorId);

        var parsed = Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, cellAnchorId);

        var cell = target.Resolve(_doc!);
        if (cell is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", cellAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var p in cell.Elements(W.p).ToList()) p.Remove();

            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                UnidHelper.AssignToSelfAndDescendants(p);
                cell.Add(p);
                PromoteHyperlinkRelationships(p);
            }
            // A table cell must contain at least one paragraph per OOXML schema.
            if (!cell.Elements(W.p).Any())
                cell.Add(new XElement(W.p));

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }

    // ─── Tier C: formatting ──────────────────────────────────────────────

    public EditResult ApplyFormat(string anchorId, CharSpan? span, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (op is null) return EditResult.Fail(EditErrorCode.MalformedMarkdown, "null format op", anchorId);
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ApplyFormat requires a paragraph anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        var totalText = ParagraphText(element);
        var actualSpan = span ?? new CharSpan(0, totalText.Length);
        if (actualSpan.Start < 0 || actualSpan.Length < 0 ||
            actualSpan.Start + actualSpan.Length > totalText.Length)
            return EditResult.Fail(EditErrorCode.OffsetOutOfRange,
                $"span [{actualSpan.Start},{actualSpan.Start + actualSpan.Length}) out of [0,{totalText.Length})", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            SplitRunsAtOffset(element, actualSpan.Start);
            SplitRunsAtOffset(element, actualSpan.Start + actualSpan.Length);

            int consumed = 0;
            foreach (var run in InlineRuns(element).ToList())
            {
                var runText = RunText(run);
                int runStart = consumed;
                int runEnd = consumed + runText.Length;
                consumed = runEnd;
                if (runEnd <= actualSpan.Start || runStart >= actualSpan.Start + actualSpan.Length) continue;
                ApplyFormatToRun(run, op);
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult SetParagraphStyle(string anchorId, string styleId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetParagraphStyle requires a paragraph anchor", anchorId);

        var stylesPart = _doc!.MainDocumentPart!.StyleDefinitionsPart;
        var stylesXml = stylesPart?.GetXDocument().Root;
        bool exists = stylesXml?.Elements(W.style)
            .Any(st => (string?)st.Attribute(W.styleId) == styleId) ?? false;
        if (!exists)
            return EditResult.Fail(EditErrorCode.UnknownStyle, $"style id not found: {styleId}", anchorId);

        var element = target.Resolve(_doc);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new XElement(W.pPr); element.AddFirst(pPr); }
            pPr.Element(W.pStyle)?.Remove();
            pPr.AddFirst(new XElement(W.pStyle, new XAttribute(W.val, styleId)));

            InvalidateProjectionCache();
            // Anchor kind may have flipped (e.g., p → h); look it up in the fresh index.
            var freshIndex = Project().AnchorIndex;
            var updated = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = ProjectScope(target),
            };
        }
        catch (Exception ex)
        {
            LastInternalError = ex;
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    public EditResult SetListLevel(string anchorId, int levelDelta)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetListLevel requires a list-item anchor", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);
        var numPr = element.Element(W.pPr)?.Element(W.numPr);
        if (numPr is null)
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "no numPr on this paragraph", anchorId);

        var ilvl = numPr.Element(W.ilvl);
        int current = ilvl is null ? 0 : int.Parse((string?)ilvl.Attribute(W.val) ?? "0");
        int next = current + levelDelta;
        if (next < 0 || next > 8)
            return EditResult.Fail(EditErrorCode.InvalidListLevel,
                $"resulting list level {next} out of [0,8]", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        ilvl?.Remove();
        numPr.Add(new XElement(W.ilvl, new XAttribute(W.val, next)));
        InvalidateProjectionCache();
        return new EditResult
        {
            Success = true,
            Modified = new[] { target.Anchor },
            Patch = ProjectScope(target),
        };
    }

    public EditResult RemoveListMembership(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var target = FindAnchor(anchorId);
        if (target is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind != "li")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "RemoveListMembership requires list-item anchor", anchorId);
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        element.Element(W.pPr)?.Element(W.numPr)?.Remove();
        InvalidateProjectionCache();
        var fresh = Project().AnchorIndex;
        var updated = fresh.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;
        return new EditResult
        {
            Success = true,
            Modified = new[] { updated },
            Patch = ProjectScope(target),
        };
    }

    // ─── Undo / Redo ─────────────────────────────────────────────────────

    public bool Undo()
    {
        if (_disposed) return false;
        var (preOp, ok) = _history.PopForUndo();
        if (!ok) return false;
        _history.RecordForRedo(TakeSnapshot());
        RestoreSnapshot(preOp);
        return true;
    }

    public bool Redo()
    {
        if (_disposed) return false;
        var (postOp, ok) = _history.PopForRedo();
        if (!ok) return false;
        _history.PushBackForUndo(TakeSnapshot());
        RestoreSnapshot(postOp);
        return true;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _doc?.Dispose();
        _stream?.Dispose();
        _doc = null;
        _stream = null;
    }

    // ─── Internal mutation helpers (used by tier methods landing in later phases) ───

    internal void InvalidateProjectionCache() => _cachedProjection = null;

    internal sealed record DocumentSnapshot(XDocument MainXml);

    internal DocumentSnapshot TakeSnapshot()
    {
        var main = _doc!.MainDocumentPart!.GetXDocument();
        return new DocumentSnapshot(new XDocument(main));
    }

    internal void RestoreSnapshot(DocumentSnapshot snapshot)
    {
        var part = _doc!.MainDocumentPart!;
        part.PutXDocument(new XDocument(snapshot.MainXml));
        InvalidateProjectionCache();
    }

    internal int NextRevisionId() => System.Threading.Interlocked.Increment(ref _revisionCounter);

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocxSession));
    }

    private static string ElementTextPreview(XElement element)
    {
        var text = string.Concat(element.Descendants(W.t).Select(t => (string)t));
        return text.Length > 80 ? text.Substring(0, 80) + "…" : text;
    }

    // ─── Mutation helpers (shared across tiers) ───────────────────────────

    internal MarkdownPatch ProjectScope(AnchorTarget target)
    {
        // Phase 3 implementation: re-project the whole document. The patch contract
        // (smallest enclosing block) is honored by ScopeAnchorId; the markdown payload
        // is the full projection until we optimize this in a later phase.
        var fresh = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
        return new MarkdownPatch(target.Anchor.Id, fresh.Markdown);
    }

    // Zero-width, semantically-significant inline markers that must survive ReplaceText.
    // Discarding them silently destroys bookmark/comment/permission ranges that point
    // into the paragraph from other parts of the document.
    private static readonly HashSet<XName> PreservedMarkerNames = new()
    {
        W.bookmarkStart, W.bookmarkEnd,
        W.commentRangeStart, W.commentRangeEnd, W.commentReference,
        W.permStart, W.permEnd,
        W.proofErr,
    };

    private static (List<XElement> pre, List<XElement> post) ExtractWrappingMarkers(XElement paragraph)
    {
        var children = paragraph.Elements().Where(e => e.Name != W.pPr).ToList();
        int firstRunIdx = children.FindIndex(IsInlineChild);
        int lastRunIdx = children.FindLastIndex(IsInlineChild);
        var pre = new List<XElement>();
        var post = new List<XElement>();
        for (int i = 0; i < children.Count; i++)
        {
            var c = children[i];
            if (!PreservedMarkerNames.Contains(c.Name)) continue;
            if (firstRunIdx < 0 || i < firstRunIdx) pre.Add(c);
            else if (i > lastRunIdx) post.Add(c);
            else pre.Add(c); // interleaved → wrap from the start (best-effort)
        }
        return (pre, post);
    }

    /// <summary>
    /// If <paramref name="paragraph"/> carries a resolvable <c>w:numPr</c> auto-number
    /// (e.g. <c>"1."</c>, <c>"Fourth"</c>), strip a matching leading prefix from
    /// <paramref name="payload"/> plus one optional separator character (ASCII space,
    /// tab, or NBSP — matching the projector's emission and the common variants an
    /// agent might use). Idempotent when the prefix isn't present.
    /// </summary>
    private string StripResolvedAutoNumberPrefix(XElement paragraph, string payload)
    {
        if (string.IsNullOrEmpty(payload)) return payload;
        // ListItemRetrieverSettings is internal to the projector; pass null so the
        // resolver uses defaults that match what the projector itself emits.
        var prefix = Internal.ListNumberResolver.Resolve(paragraph, _doc!);
        if (string.IsNullOrEmpty(prefix)) return payload;
        if (!payload.StartsWith(prefix, StringComparison.Ordinal)) return payload;

        var after = payload.Substring(prefix.Length);
        if (after.Length > 0 && (after[0] == ' ' || after[0] == '\t' || after[0] == ' '))
            after = after.Substring(1);
        return after;
    }

    private static void ApplyReplaceTextAccept(XElement paragraph, IReadOnlyList<Internal.ParsedBlock> blocks)
    {
        var pPr = paragraph.Element(W.pPr);
        var (preMarkers, postMarkers) = ExtractWrappingMarkers(paragraph);
        paragraph.RemoveNodes();
        if (pPr is not null) paragraph.Add(pPr);
        foreach (var m in preMarkers) paragraph.Add(m);
        if (blocks.Count > 0)
            foreach (var run in blocks[0].RunElements)
                paragraph.Add(new XElement(run));
        foreach (var m in postMarkers) paragraph.Add(m);
    }

    private void ApplyReplaceTextTracked(XElement paragraph, IReadOnlyList<Internal.ParsedBlock> blocks)
    {
        var author = _settings.RevisionAuthor ?? "docxodus";
        var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");

        // Wrap existing runs in w:del (converting w:t to w:delText).
        var existingRuns = paragraph.Elements(W.r).ToList();
        if (existingRuns.Count > 0)
        {
            var del = new XElement(W.del,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date));
            foreach (var run in existingRuns)
            {
                run.Remove();
                foreach (var t in run.Elements(W.t).ToList())
                {
                    var dt = new XElement(W.delText,
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        (string)t);
                    t.ReplaceWith(dt);
                }
                del.Add(run);
            }
            paragraph.Add(del);
        }

        if (blocks.Count > 0 && blocks[0].RunElements.Count > 0)
        {
            var ins = new XElement(W.ins,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date));
            foreach (var run in blocks[0].RunElements)
                ins.Add(new XElement(run));
            paragraph.Add(ins);
        }
    }

    private void WrapRunsInDel(XElement element)
    {
        var author = _settings.RevisionAuthor ?? "docxodus";
        var date = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ");
        foreach (var run in element.Elements(W.r).ToList())
        {
            run.Remove();
            foreach (var t in run.Elements(W.t).ToList())
                t.ReplaceWith(new XElement(W.delText,
                    new XAttribute(XNamespace.Xml + "space", "preserve"),
                    (string)t));
            var del = new XElement(W.del,
                new XAttribute(W.id, NextRevisionId()),
                new XAttribute(W.author, author),
                new XAttribute(W.date, date),
                run);
            element.Add(del);
        }
    }

    private void PromoteHyperlinkRelationships(XElement paragraph)
    {
        var main = _doc!.MainDocumentPart!;
        // Reuse an existing relationship when the same URL has already been registered.
        // Without dedup, every ReplaceText with a link adds a fresh rId; an agent loop
        // that edits the same paragraph N times grows the .rels file unboundedly.
        var existing = main.HyperlinkRelationships
            .GroupBy(rl => rl.Uri.ToString())
            .ToDictionary(g => g.Key, g => g.First().Id);
        foreach (var link in paragraph.Descendants(W.hyperlink).ToList())
        {
            var hrefAttr = link.Attribute(Internal.MarkdownPayloadParser.HrefAttr);
            if (hrefAttr is null) continue;
            var url = hrefAttr.Value;
            string relId;
            if (existing.TryGetValue(url, out var foundId)) relId = foundId;
            else
            {
                var rel = main.AddHyperlinkRelationship(
                    new Uri(url, UriKind.RelativeOrAbsolute), true);
                relId = rel.Id;
                existing[url] = relId;
            }
            link.SetAttributeValue(R.id, relId);
            hrefAttr.Remove();
        }
    }

    private static void ApplyFormatToRun(XElement run, FormatOp op)
    {
        var rPr = run.Element(W.rPr);
        if (rPr is null) { rPr = new XElement(W.rPr); run.AddFirst(rPr); }

        static void Toggle(XElement rPr, XName name, bool? set)
        {
            if (set is null) return;
            var existing = rPr.Element(name);
            if (set.Value && existing is null) rPr.Add(new XElement(name));
            else if (!set.Value) existing?.Remove();
        }

        Toggle(rPr, W.b, op.Bold);
        Toggle(rPr, W.i, op.Italic);
        Toggle(rPr, W.strike, op.Strike);

        if (op.Underline is true)
        {
            rPr.Element(W.u)?.Remove();
            rPr.Add(new XElement(W.u, new XAttribute(W.val, "single")));
        }
        else if (op.Underline is false) rPr.Element(W.u)?.Remove();

        if (op.Code is true)
        {
            rPr.Element(W.rStyle)?.Remove();
            rPr.Add(new XElement(W.rStyle, new XAttribute(W.val, "Code")));
        }
        else if (op.Code is false) rPr.Element(W.rStyle)?.Remove();

        if (op.Color is not null)
        {
            rPr.Element(W.color)?.Remove();
            if (op.Color.Length > 0)
                rPr.Add(new XElement(W.color, new XAttribute(W.val, op.Color)));
        }

        if (op.RunStyle is not null)
        {
            rPr.Element(W.rStyle)?.Remove();
            if (op.RunStyle.Length > 0)
                rPr.Add(new XElement(W.rStyle, new XAttribute(W.val, op.RunStyle)));
        }
    }

    internal static XElement BuildParagraphFromParsedBlock(Internal.ParsedBlock block)
    {
        var p = new XElement(W.p);
        var pPr = new XElement(W.pPr);

        switch (block.Kind)
        {
            case Internal.ParserBlockKind.Heading1:
            case Internal.ParserBlockKind.Heading2:
            case Internal.ParserBlockKind.Heading3:
            case Internal.ParserBlockKind.Heading4:
            case Internal.ParserBlockKind.Heading5:
            case Internal.ParserBlockKind.Heading6:
                {
                    int level = (int)block.Kind - (int)Internal.ParserBlockKind.Heading1 + 1;
                    pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, $"Heading{level}")));
                    break;
                }
            case Internal.ParserBlockKind.Quote:
                pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, "Quote")));
                break;
            case Internal.ParserBlockKind.Code:
                pPr.Add(new XElement(W.pStyle, new XAttribute(W.val, "Code")));
                break;
            // List items: numPr inheritance not auto-injected in v1 — caller can use
            // SetListLevel afterwards if needed. The bare paragraph will project as a
            // normal paragraph until numbering is added.
        }

        if (pPr.HasElements) p.Add(pPr);
        foreach (var run in block.RunElements)
            p.Add(new XElement(run));
        return p;
    }

    internal static string ParserBlockKindToAnchorKind(Internal.ParserBlockKind kind) => kind switch
    {
        Internal.ParserBlockKind.Heading1
            or Internal.ParserBlockKind.Heading2
            or Internal.ParserBlockKind.Heading3
            or Internal.ParserBlockKind.Heading4
            or Internal.ParserBlockKind.Heading5
            or Internal.ParserBlockKind.Heading6 => "h",
        Internal.ParserBlockKind.BulletItem
            or Internal.ParserBlockKind.OrderedItem => "li",
        _ => "p",
    };

    /// <summary>
    /// Mirror the classifier used by <see cref="WmlToMarkdownConverter"/> so the kind
    /// reported in <see cref="EditResult.Created"/> matches what the projector will
    /// emit on the next <see cref="DocxSession.Project"/>. If we used the parser's
    /// kind blindly, a bullet-payload paragraph without a <c>w:numPr</c> would be
    /// reported as "li" but appear as "p" in the projection — a stale anchor id.
    /// </summary>
    internal static string ClassifyParagraphKind(XElement paragraph)
    {
        var pPr = paragraph.Element(W.pPr);
        var styleId = (string?)pPr?.Element(W.pStyle)?.Attribute(W.val);
        if (!string.IsNullOrEmpty(styleId)
            && (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
                || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase)))
            return "h";
        if (pPr?.Element(W.numPr) is not null) return "li";
        return "p";
    }

    /// <summary>
    /// Copy <c>w:numPr</c> from a nearby sibling list item into the new paragraph so
    /// a bullet/ordered-item payload actually renders as part of an existing list.
    /// Walks previous siblings first (closest match first), then next siblings.
    /// No-op when no sibling carries numbering — caller then reports kind="p" via
    /// <see cref="ClassifyParagraphKind"/>.
    /// </summary>
    private static void TryInheritNumPrFromSibling(XElement newParagraph, XElement anchorElement)
    {
        XElement? donorNumPr = null;
        XElement? donorPStyle = null;
        foreach (var sib in anchorElement.ElementsBeforeSelf().Reverse()
                                .Concat(new[] { anchorElement })
                                .Concat(anchorElement.ElementsAfterSelf()))
        {
            if (sib.Name != W.p) continue;
            var nump = sib.Element(W.pPr)?.Element(W.numPr);
            if (nump is null) continue;
            donorNumPr = nump;
            donorPStyle = sib.Element(W.pPr)?.Element(W.pStyle);
            break;
        }
        if (donorNumPr is null) return;

        var pPr = newParagraph.Element(W.pPr);
        if (pPr is null) { pPr = new XElement(W.pPr); newParagraph.AddFirst(pPr); }
        if (pPr.Element(W.numPr) is null) pPr.Add(new XElement(donorNumPr));
        if (donorPStyle is not null && pPr.Element(W.pStyle) is null)
            pPr.AddFirst(new XElement(donorPStyle));
    }

    // Top-level inline children of <w:p> that participate in text flow.
    // Hyperlinks, sdts, fldSimple and smartTag are transparent containers — their
    // descendant runs contribute to the paragraph's visible text. Bookmark/comment
    // markers (zero-width) are tracked separately and not enumerated here.
    private static readonly HashSet<XName> InlineContainerNames = new()
    {
        W.hyperlink, W.sdt, W.fldSimple, W.smartTag,
    };

    private static bool IsInlineChild(XElement e) =>
        e.Name == W.r || InlineContainerNames.Contains(e.Name);

    /// <summary>
    /// All <c>&lt;w:r&gt;</c> elements that contribute to the paragraph's visible text,
    /// in document order — including runs nested inside hyperlinks, sdts, fldSimple,
    /// smartTags. Iterating only <c>Elements(W.r)</c> silently skips hyperlink-internal
    /// runs, which produced the bugs documented in DS080-DS090.
    /// </summary>
    internal static IEnumerable<XElement> InlineRuns(XElement paragraph)
    {
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == W.r) yield return child;
            else if (InlineContainerNames.Contains(child.Name))
                foreach (var run in child.Descendants(W.r))
                    yield return run;
        }
    }

    internal static string ParagraphText(XElement paragraph) =>
        string.Concat(InlineRuns(paragraph).Select(RunText));

    internal static string RunText(XElement run) =>
        string.Concat(run.Elements(W.t).Select(t => (string)t));

    private static int InlineChildTextLength(XElement child) =>
        string.Concat(child.DescendantsAndSelf(W.t).Select(t => (string)t)).Length;

    /// <summary>
    /// If a run straddles <paramref name="offset"/>, split it into two adjacent runs
    /// at that offset. Walks runs inside hyperlinks/sdts/etc. too, so the boundary
    /// is clean regardless of which container the run lives in. The new sibling run
    /// is inserted into the same parent as the original (preserving hyperlink/sdt
    /// membership for the keep-half).
    /// </summary>
    internal static void SplitRunsAtOffset(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var run in InlineRuns(paragraph).ToList())
        {
            var runText = RunText(run);
            if (consumed == offset) return;
            if (consumed + runText.Length <= offset) { consumed += runText.Length; continue; }
            int splitAt = offset - consumed;
            if (splitAt <= 0) return;

            var keep = runText.Substring(0, splitAt);
            var move = runText.Substring(splitAt);

            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"), keep));

            var rPr = run.Element(W.rPr);
            var newRun = new XElement(W.r);
            if (rPr is not null) newRun.Add(new XElement(rPr));
            newRun.Add(new XElement(W.t,
                new XAttribute(XNamespace.Xml + "space", "preserve"), move));
            run.AddAfterSelf(newRun);
            return;
        }
    }

    /// <summary>
    /// Ensures no top-level inline child straddles <paramref name="offset"/>: if a
    /// hyperlink (or other splittable container) crosses the boundary, it's split
    /// into two sibling containers sharing the same attributes (e.g. <c>r:id</c>),
    /// each holding half the runs. After this call, <see cref="MoveInlineChildrenAfter"/>
    /// can move whole-child elements without slicing through anything.
    /// </summary>
    internal static void SplitInlineContainersAtOffset(XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var child in paragraph.Elements().Where(IsInlineChild).ToList())
        {
            int len = InlineChildTextLength(child);
            if (consumed + len <= offset) { consumed += len; continue; }
            if (consumed == offset) return; // boundary already clean
            int local = offset - consumed;

            if (child.Name == W.hyperlink)
                SplitHyperlinkAt(child, local);
            // For <w:r>: SplitRunsAtOffset already handled it. For sdt/fldSimple/smartTag:
            // treat as atomic — splitting these requires semantic care; the whole element
            // stays with whichever side its leading run lands on.
            return;
        }
    }

    private static void SplitHyperlinkAt(XElement hyperlink, int localOffset)
    {
        // Split runs inside the hyperlink at the local offset (works because SplitRunsAtOffset
        // walks descendants through container types).
        SplitRunsAtOffset(hyperlink, localOffset);

        int consumed = 0;
        var movedRuns = new List<XElement>();
        foreach (var run in hyperlink.Elements(W.r).ToList())
        {
            int len = RunText(run).Length;
            if (consumed >= localOffset) movedRuns.Add(run);
            consumed += len;
        }
        if (movedRuns.Count == 0) return;

        var newLink = new XElement(W.hyperlink);
        foreach (var a in hyperlink.Attributes()) newLink.SetAttributeValue(a.Name, a.Value);
        foreach (var run in movedRuns) { run.Remove(); newLink.Add(run); }
        hyperlink.AddAfterSelf(newLink);
    }

    /// <summary>
    /// Move every paragraph child (inline run/container OR zero-width marker)
    /// whose position is at or past <paramref name="offset"/> from
    /// <paramref name="paragraph"/> into <paramref name="destination"/>. Inline
    /// children advance the position counter by their text length; markers
    /// (bookmarkStart/End, comment range markers, etc.) advance it by 0 and so
    /// inherit the position they're sandwiched between.
    /// </summary>
    internal static void MoveInlineChildrenAfter(XElement paragraph, int offset, XElement destination)
    {
        int consumed = 0;
        var toMove = new List<XElement>();
        foreach (var child in paragraph.Elements().ToList())
        {
            if (child.Name == W.pPr) continue;
            int len = IsInlineChild(child) ? InlineChildTextLength(child) : 0;
            if (consumed >= offset) toMove.Add(child);
            consumed += len;
        }
        foreach (var c in toMove) { c.Remove(); destination.Add(c); }
    }
}
