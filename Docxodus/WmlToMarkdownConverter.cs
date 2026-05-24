#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
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
        throw new NotImplementedException(
            "Anchor resolution is part of the scaffolded WmlToMarkdownConverter. " +
            "See docs/architecture/markdown_projection.md for the design.");
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
    /// Convert a <see cref="WmlDocument"/> to its markdown projection.
    /// </summary>
    public static MarkdownProjection Convert(WmlDocument document, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(settings);

        throw new NotImplementedException(
            "WmlToMarkdownConverter is scaffolded — projection logic ships in phases. " +
            "See docs/architecture/markdown_projection.md.");
    }

    /// <summary>
    /// Convert an open <see cref="WordprocessingDocument"/> to its markdown projection
    /// without round-tripping through <see cref="WmlDocument"/>.
    /// </summary>
    public static MarkdownProjection Convert(WordprocessingDocument document, WmlToMarkdownConverterSettings settings)
    {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(settings);

        throw new NotImplementedException(
            "WmlToMarkdownConverter is scaffolded — projection logic ships in phases. " +
            "See docs/architecture/markdown_projection.md.");
    }
}
