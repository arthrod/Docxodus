#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Pure resolvers for the block-metadata read surface — no mutation, no undo
/// snapshots. Each method takes a live <see cref="WordprocessingDocument"/>
/// and an <see cref="AnchorTarget"/> and walks the OOXML tree to assemble
/// the requested record.
/// </summary>
internal static class BlockMetadataOps
{
    /// <summary>Resolve <see cref="BlockMetadata"/> for the given anchor target.</summary>
    public static BlockMetadata? GetBlockMetadata(WordprocessingDocument doc, AnchorTarget target)
    {
        var element = target.Resolve(doc);
        if (element is null) return null;

        var styleId = ResolveStyleId(element);
        var styleName = styleId is null ? null : ResolveStyleName(doc, styleId);
        var outlineLevel = ResolveOutlineLevel(element, styleId);
        var list = ResolveListMembership(doc, element, target);
        var hasFormatting = HasInlineFormatting(element);

        return new BlockMetadata
        {
            AnchorId = target.Anchor.Id,
            Kind = target.Anchor.Kind,
            Scope = target.Anchor.Scope,
            StyleId = styleId,
            StyleName = styleName,
            OutlineLevel = outlineLevel,
            List = list,
            HasInlineFormatting = hasFormatting,
        };
    }

    /// <summary>Resolve <see cref="ListMembership"/> for the given target, or null if not a list item.</summary>
    public static ListMembership? GetListMembership(WordprocessingDocument doc, AnchorTarget target)
    {
        var element = target.Resolve(doc);
        return element is null ? null : ResolveListMembership(doc, element, target);
    }

    /// <summary>Resolve <see cref="SectionInfo"/> for the given target, or null for non-body anchors.</summary>
    public static SectionInfo? GetSectionInfo(WordprocessingDocument doc, AnchorTarget target)
    {
        // Implemented in Task 7.
        return null;
    }

    // ─── Internals ──────────────────────────────────────────────────────

    private static string? ResolveStyleId(XElement element)
    {
        if (element.Name == W.p)
            return (string?)element.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
        if (element.Name == W.tbl)
            return (string?)element.Element(W.tblPr)?.Element(W.tblStyle)?.Attribute(W.val);
        return null;
    }

    private static string? ResolveStyleName(WordprocessingDocument doc, string styleId)
    {
        var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
        var root = stylesPart?.GetXDocument().Root;
        if (root is null) return null;
        var style = root.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == styleId);
        return (string?)style?.Element(W.name)?.Attribute(W.val);
    }

    private static int? ResolveOutlineLevel(XElement element, string? styleId)
    {
        if (element.Name != W.p) return null;

        var outlineLvl = element.Element(W.pPr)?.Element(W.outlineLvl);
        if (outlineLvl is not null && int.TryParse((string?)outlineLvl.Attribute(W.val), out var lvl))
            return lvl;

        // Infer from Heading1..Heading9 style id.
        if (styleId is { Length: > 7 }
            && styleId.StartsWith("Heading", System.StringComparison.Ordinal)
            && int.TryParse(styleId.AsSpan(7), out var hLvl)
            && hLvl >= 1 && hLvl <= 9)
        {
            return hLvl - 1;  // outlineLvl is 0-based
        }
        return null;
    }

    private static ListMembership? ResolveListMembership(
        WordprocessingDocument doc, XElement element, AnchorTarget target)
    {
        if (element.Name != W.p) return null;

        // Inline w:numPr beats style-inherited numPr.
        var pPr = element.Element(W.pPr);
        var numPr = pPr?.Element(W.numPr);
        bool fromStyle = false;

        if (numPr is null)
        {
            // Walk the pStyle chain looking for a style that contributes numPr.
            var styleId = (string?)pPr?.Element(W.pStyle)?.Attribute(W.val);
            numPr = ResolveStyleNumPr(doc, styleId);
            fromStyle = numPr is not null;
        }

        if (numPr is null) return null;

        var numIdEl = numPr.Element(W.numId);
        if (numIdEl is null) return null;
        if (!int.TryParse((string?)numIdEl.Attribute(W.val), out var numId)) return null;

        var ilvlEl = numPr.Element(W.ilvl);
        int level = 0;
        if (ilvlEl is not null) int.TryParse((string?)ilvlEl.Attribute(W.val), out level);

        var numberingPart = doc.MainDocumentPart?.NumberingDefinitionsPart;
        var numberingRoot = numberingPart?.GetXDocument().Root;
        if (numberingRoot is null) return null;

        var numEl = numberingRoot.Elements(W.num)
            .FirstOrDefault(n => (string?)n.Attribute(W.numId) == numId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (numEl is null) return null;

        var abstractNumIdEl = numEl.Element(W.abstractNumId);
        if (abstractNumIdEl is null) return null;
        if (!int.TryParse((string?)abstractNumIdEl.Attribute(W.val), out var abstractNumId)) return null;

        var abstractNumEl = numberingRoot.Elements(W.abstractNum)
            .FirstOrDefault(a => (string?)a.Attribute(W.abstractNumId) == abstractNumId.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (abstractNumEl is null) return null;

        var lvlEl = abstractNumEl.Elements(W.lvl)
            .FirstOrDefault(l => (string?)l.Attribute(W.ilvl) == level.ToString(System.Globalization.CultureInfo.InvariantCulture));
        var format = ParseNumberFormat((string?)lvlEl?.Element(W.numFmt)?.Attribute(W.val));

        // Start override from the w:num's lvlOverride for this level.
        int? startOverride = null;
        var lvlOverrideEl = numEl.Elements(W.lvlOverride)
            .FirstOrDefault(o => (string?)o.Attribute(W.ilvl) == level.ToString(System.Globalization.CultureInfo.InvariantCulture));
        if (lvlOverrideEl is not null)
        {
            var startOverrideEl = lvlOverrideEl.Element(W.startOverride);
            if (startOverrideEl is not null && int.TryParse((string?)startOverrideEl.Attribute(W.val), out var so))
                startOverride = so;
        }

        return new ListMembership
        {
            NumId = numId,
            AbstractNumId = abstractNumId,
            Level = level,
            Format = format,
            StartOverride = startOverride,
            IsAutoNumbered = true,
            FromStyle = fromStyle,
            GeneratedLabel = target.AutoNumberPrefix,
        };
    }

    private static XElement? ResolveStyleNumPr(WordprocessingDocument doc, string? styleId)
    {
        if (string.IsNullOrEmpty(styleId)) return null;
        var stylesPart = doc.MainDocumentPart?.StyleDefinitionsPart;
        var root = stylesPart?.GetXDocument().Root;
        if (root is null) return null;

        // Walk style → basedOn → ... up to 16 levels (cycle-guard depth).
        var visited = new HashSet<string>(System.StringComparer.Ordinal);
        var current = styleId;
        for (int i = 0; i < 16 && current is not null; i++)
        {
            if (!visited.Add(current)) return null;  // cycle
            var style = root.Elements(W.style).FirstOrDefault(s => (string?)s.Attribute(W.styleId) == current);
            if (style is null) return null;
            var numPr = style.Element(W.pPr)?.Element(W.numPr);
            if (numPr is not null) return numPr;
            current = (string?)style.Element(W.basedOn)?.Attribute(W.val);
        }
        return null;
    }

    private static NumberFormat ParseNumberFormat(string? raw) => raw switch
    {
        "bullet" => NumberFormat.Bullet,
        "upperLetter" => NumberFormat.UpperLetter,
        "lowerLetter" => NumberFormat.LowerLetter,
        "upperRoman" => NumberFormat.UpperRoman,
        "lowerRoman" => NumberFormat.LowerRoman,
        _ => NumberFormat.Decimal,  // includes "decimal" and any unrecognized format
    };

    private static bool HasInlineFormatting(XElement element)
    {
        return element.Descendants(W.r).Any(r =>
        {
            var rPr = r.Element(W.rPr);
            return rPr is not null && rPr.Elements().Any();
        });
    }
}
