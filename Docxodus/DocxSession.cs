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
