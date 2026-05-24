#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

/// <summary>
/// Which parts of the OOXML package to include in the markdown projection.
/// </summary>
[Flags]
public enum ProjectionScopes
{
    Body = 1,
    Headers = 2,
    Footers = 4,
    Footnotes = 8,
    Endnotes = 16,
    Comments = 32,
    All = Body | Headers | Footers | Footnotes | Endnotes | Comments
}

/// <summary>
/// How anchor markers are rendered in the markdown output.
/// </summary>
public enum AnchorRenderMode
{
    /// <summary>Anchor appears on its own line before each block element (default).</summary>
    Block,
    /// <summary>Block anchors plus inline {#...} markers for spans (comments, hyperlinks, fields).</summary>
    BlockAndInline,
    /// <summary>No anchor markers in the output (projection only, no addressing).</summary>
    None
}

/// <summary>
/// Strategy for converting <c>w:tbl</c> elements that don't fit GFM pipe-table constraints.
/// </summary>
public enum TableRenderMode
{
    /// <summary>Emit GFM pipe tables when possible, opaque anchor blocks otherwise (default).</summary>
    GfmWithOpaqueFallback,
    /// <summary>Always emit GFM pipe tables, flattening complex structure with possible loss.</summary>
    AlwaysGfm,
    /// <summary>Always emit opaque anchor blocks. Useful when callers will fetch tables via the SDK.</summary>
    AlwaysOpaque
}

/// <summary>
/// How tracked changes are handled when projecting to markdown.
/// </summary>
public enum TrackedChangeMode
{
    /// <summary>Accept all revisions before conversion (default).</summary>
    Accept,
    /// <summary>Render insertions and deletions inline as <c>{+ins+}</c> / <c>{-del-}</c>.</summary>
    RenderInline,
    /// <summary>Accept insertions, drop deletions.</summary>
    StripDeletions
}

/// <summary>
/// Settings controlling markdown projection of a Word document. See
/// <c>docs/architecture/markdown_projection.md</c> for the full specification.
/// </summary>
public class WmlToMarkdownConverterSettings
{
    /// <summary>Which package parts to include in the projection.</summary>
    public ProjectionScopes Scopes { get; set; } = ProjectionScopes.All;

    /// <summary>
    /// Offset added to Word heading levels when emitting markdown headings.
    /// Example: an offset of 1 turns Word Heading1 into a markdown <c>##</c>.
    /// </summary>
    public int HeadingLevelOffset { get; set; } = 0;

    /// <summary>How anchor markers are emitted in the output.</summary>
    public AnchorRenderMode AnchorMode { get; set; } = AnchorRenderMode.Block;

    /// <summary>Strategy for tables that don't fit GFM pipe-table syntax.</summary>
    public TableRenderMode TableMode { get; set; } = TableRenderMode.GfmWithOpaqueFallback;

    /// <summary>
    /// Maximum characters per cell before a simple table is downgraded to an opaque
    /// anchor block. Ignored when <see cref="TableMode"/> is <see cref="TableRenderMode.AlwaysGfm"/>
    /// or <see cref="TableRenderMode.AlwaysOpaque"/>.
    /// </summary>
    public int TableInlineCellMax { get; set; } = 80;

    /// <summary>How tracked changes (<c>w:ins</c>, <c>w:del</c>) are projected.</summary>
    public TrackedChangeMode TrackedChanges { get; set; } = TrackedChangeMode.Accept;

    /// <summary>
    /// Resolve list numbering (<c>w:numPr</c>) to literal markers (<c>1.</c>, <c>a.</c>, etc.).
    /// When false, list items are emitted as plain markdown unordered items.
    /// </summary>
    public bool ResolveNumbering { get; set; } = true;

    /// <summary>
    /// Builds the URI used for image references in the output. Defaults to
    /// <c>docxodus://img/{unid}</c>; callers that want data URIs or HTTP URLs can override.
    /// </summary>
    public Func<ImageInfo, string>? ImageUriBuilder { get; set; }
}

/// <summary>
/// Identifies a single addressable element in the markdown projection.
/// Anchor ids have the form <c>kind:scope:unid</c> (e.g. <c>p:body:a1b2c3d4</c>).
/// </summary>
public readonly record struct Anchor(string Id, string Kind, string Scope, string Unid)
{
    /// <summary>The block-level token rendered in the projection (e.g. <c>{#p:body:a1b2c3d4}</c>).</summary>
    public string Token => $"{{#{Id}}}";
}

/// <summary>
/// Resolved location of an anchor in the underlying OOXML package — sufficient to walk
/// back to the source <see cref="XElement"/> and apply edits via the Open XML SDK.
/// </summary>
public sealed class AnchorTarget
{
    /// <summary>The anchor this target resolves.</summary>
    required public Anchor Anchor { get; init; }

    /// <summary>URI of the package part containing the element (e.g. main document, header part).</summary>
    required public string PartUri { get; init; }

    /// <summary>Stable Unid of the element. Use with Docxodus' Unid lookup to fetch the live XElement.</summary>
    required public string Unid { get; init; }

    /// <summary>
    /// Resolves this anchor to its current <see cref="XElement"/> inside the given document.
    /// Returns null if the element has been removed since projection.
    /// </summary>
    public XElement? Resolve(WordprocessingDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);

        var main = document.MainDocumentPart;
        if (main == null) return null;

        OpenXmlPart? part = null;
        var targetUri = PartUri;
        if (main.Uri.ToString() == targetUri) part = main;
        else
        {
            foreach (var h in main.HeaderParts) if (h.Uri.ToString() == targetUri) { part = h; break; }
            if (part == null)
                foreach (var f in main.FooterParts) if (f.Uri.ToString() == targetUri) { part = f; break; }
            if (part == null && main.FootnotesPart?.Uri.ToString() == targetUri) part = main.FootnotesPart;
            if (part == null && main.EndnotesPart?.Uri.ToString() == targetUri) part = main.EndnotesPart;
            if (part == null && main.WordprocessingCommentsPart?.Uri.ToString() == targetUri) part = main.WordprocessingCommentsPart;
        }

        var root = part?.GetXDocument().Root;
        return root?.DescendantsAndSelf()
            .FirstOrDefault(e => (string?)e.Attribute(PtOpenXml.Unid) == Unid);
    }
}

/// <summary>
/// The output of a markdown projection: the rendered text plus the anchor index that maps
/// each <c>{#...}</c> token back to a location in the OOXML package.
/// </summary>
public sealed class MarkdownProjection
{
    required public string Markdown { get; init; }
    required public IReadOnlyDictionary<string, AnchorTarget> AnchorIndex { get; init; }
}

public partial class WmlDocument
{
    /// <summary>
    /// Project this document to anchor-addressed markdown. See
    /// <c>docs/architecture/markdown_projection.md</c> for the projection spec.
    /// </summary>
    public MarkdownProjection ConvertToMarkdown(WmlToMarkdownConverterSettings settings)
    {
        return WmlToMarkdownConverter.Convert(this, settings);
    }
}

/// <summary>
/// Projects a Word document to anchor-addressed markdown — a stable, deterministic text view
/// suitable for tools that want to read, search, and edit Word documents the way they would
/// source files (LLM editing pipelines, structured search indexers, diff/review UIs).
///
/// This is a scaffold. The public surface is fixed; the implementation is staged in phases
/// described in <c>docs/architecture/markdown_projection.md</c>.
/// </summary>
public static class WmlToMarkdownConverter
{
    /// <summary>
    /// Convert a <see cref="WmlDocument"/> to its markdown projection. Persists any newly
    /// assigned <c>PtOpenXml.Unid</c> attributes back into <paramref name="document"/>'s
    /// underlying byte array so subsequent <see cref="AnchorTarget.Resolve"/> calls against
    /// the same bytes succeed.
    /// </summary>
    public static MarkdownProjection Convert(WmlDocument document, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(settings);

        using var stream = new OpenXmlMemoryStreamDocument(document);
        MarkdownProjection projection;
        using (var wdoc = stream.GetWordprocessingDocument())
        {
            projection = Convert(wdoc, settings);
        }
        var modified = stream.GetModifiedWmlDocument();
        document.DocumentByteArray = modified.DocumentByteArray;
        return projection;
    }

    /// <summary>
    /// Convert an open <see cref="WordprocessingDocument"/> to its markdown projection
    /// without round-tripping through <see cref="WmlDocument"/>. Mutates the in-memory
    /// XDocument of each in-scope part to add missing <c>PtOpenXml.Unid</c> attributes and
    /// persists those mutations via <c>PutXDocument</c>; the caller is responsible for
    /// saving the package itself.
    /// </summary>
    public static MarkdownProjection Convert(WordprocessingDocument document, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(settings);

        var (index, scopes) = BuildAnchorIndex(document, settings);
        var markdown = EmitMarkdown(document, settings, scopes);
        return new MarkdownProjection { Markdown = markdown, AnchorIndex = index };
    }

    // ------------------------------------------------------------------
    // Phase 1: anchor index
    // ------------------------------------------------------------------

    private sealed class ScopeInfo
    {
        required public string Name { get; init; }
        required public OpenXmlPart Part { get; init; }
        required public XElement Root { get; init; }
    }

    private static (IReadOnlyDictionary<string, AnchorTarget> Index, List<ScopeInfo> Scopes)
        BuildAnchorIndex(WordprocessingDocument doc, WmlToMarkdownConverterSettings settings)
    {
        var main = doc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no MainDocumentPart.");

        var scopes = new List<ScopeInfo>();
        if (settings.Scopes.HasFlag(ProjectionScopes.Body))
            scopes.Add(new ScopeInfo { Name = "body", Part = main, Root = main.GetXDocument().Root! });
        if (settings.Scopes.HasFlag(ProjectionScopes.Headers))
        {
            var i = 1;
            foreach (var hp in main.HeaderParts)
                scopes.Add(new ScopeInfo { Name = $"hdr{i++}", Part = hp, Root = hp.GetXDocument().Root! });
        }
        if (settings.Scopes.HasFlag(ProjectionScopes.Footers))
        {
            var i = 1;
            foreach (var fp in main.FooterParts)
                scopes.Add(new ScopeInfo { Name = $"ftr{i++}", Part = fp, Root = fp.GetXDocument().Root! });
        }
        if (settings.Scopes.HasFlag(ProjectionScopes.Footnotes) && main.FootnotesPart != null)
            scopes.Add(new ScopeInfo { Name = "fn", Part = main.FootnotesPart, Root = main.FootnotesPart.GetXDocument().Root! });
        if (settings.Scopes.HasFlag(ProjectionScopes.Endnotes) && main.EndnotesPart != null)
            scopes.Add(new ScopeInfo { Name = "en", Part = main.EndnotesPart, Root = main.EndnotesPart.GetXDocument().Root! });
        if (settings.Scopes.HasFlag(ProjectionScopes.Comments) && main.WordprocessingCommentsPart != null)
            scopes.Add(new ScopeInfo { Name = "cmt", Part = main.WordprocessingCommentsPart, Root = main.WordprocessingCommentsPart.GetXDocument().Root! });

        var index = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
        foreach (var scope in scopes)
        {
            UnidHelper.AssignToAllElements(scope.Root);
            foreach (var el in scope.Root.DescendantsAndSelf())
            {
                var kind = KindFor(el);
                if (kind == null) continue;
                var unid = (string?)el.Attribute(PtOpenXml.Unid);
                if (unid == null) continue;
                var id = $"{kind}:{scope.Name}:{unid}";
                if (index.ContainsKey(id)) continue;
                var anchor = new Anchor(id, kind, scope.Name, unid);
                index[id] = new AnchorTarget
                {
                    Anchor = anchor,
                    PartUri = scope.Part.Uri.ToString(),
                    Unid = unid,
                };
            }
            scope.Part.PutXDocument();
        }

        return (index, scopes);
    }

    /// <summary>
    /// Classify an element to its anchor <c>kind</c>. Returns <c>null</c> for elements that
    /// are not addressable by the projection (runs, inline children, formatting properties).
    /// </summary>
    private static string? KindFor(XElement el)
    {
        var n = el.Name;
        if (n == W.p) return IsHeading(el) ? "h" : IsListItem(el) ? "li" : "p";
        if (n == W.tbl) return "tbl";
        if (n == W.tr) return "tr";
        if (n == W.tc) return "tc";
        if (n == W.footnote) return "fn";
        if (n == W.endnote) return "en";
        if (n == W.comment) return "cmt";
        return null;
    }

    private static bool IsHeading(XElement p)
    {
        var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (string.IsNullOrEmpty(styleId)) return false;
        return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            || styleId.Equals("Title", StringComparison.OrdinalIgnoreCase)
            || styleId.Equals("Subtitle", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsListItem(XElement p)
        => p.Element(W.pPr)?.Element(W.numPr) != null;

    // ------------------------------------------------------------------
    // Markdown emission. Phase 1 produces an empty string; subsequent phases
    // append to this. Implemented in this file under Phase 2+ markers below.
    // ------------------------------------------------------------------

    private static string EmitMarkdown(
        WordprocessingDocument document,
        WmlToMarkdownConverterSettings settings,
        List<ScopeInfo> scopes)
    {
        // Placeholder; phases 2-7 fill this in.
        return string.Empty;
    }
}
