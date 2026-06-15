#nullable enable

using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// Synthesizes reusable bullet / decimal numbering definitions so a plain paragraph can be
/// promoted to a real list item. <see cref="DocxSession.ApplyListFormat"/> uses this when no
/// suitable numbering exists. Definitions are tagged with a fixed marker <c>w:nsid</c> per
/// format and resolved find-or-create, so the op is idempotent across calls, save/reopen, and
/// undo (the numbering part is not snapshotted; the paragraph's <c>w:numPr</c> is).
/// </summary>
internal static class NumberingFactory
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // Stable per-format markers (8-hex nsid values) used to find-or-create our own definition.
    private const string BulletNsid = "0D0CB001";
    private const string DecimalNsid = "0D0CD001";

    /// <summary>
    /// Ensure a bullet or decimal numbering definition exists and return a numId pointing at it.
    /// Only NumberFormat.Bullet and NumberFormat.Decimal are supported here.
    /// </summary>
    public static int EnsureNumbering(WordprocessingDocument doc, NumberFormat fmt)
    {
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("no MainDocumentPart");
        var part = main.NumberingDefinitionsPart;
        if (part is null)
        {
            part = main.AddNewPart<NumberingDefinitionsPart>();
            part.PutXDocument(new XDocument(
                new XElement(W + "numbering", new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName))));
        }

        var root = part.GetXDocument().Root!;
        bool bullet = fmt == NumberFormat.Bullet;
        string nsid = bullet ? BulletNsid : DecimalNsid;

        // Find our previously-synthesized abstractNum (by marker nsid), or build one.
        var abstractNum = root.Elements(W + "abstractNum")
            .FirstOrDefault(a => (string?)a.Element(W + "nsid")?.Attribute(W + "val") == nsid);
        if (abstractNum is null)
        {
            int absId = NextId(root, "abstractNum", "abstractNumId");
            abstractNum = BuildAbstractNum(bullet, absId, nsid);
            // CT_Numbering order: numPicBullet*, abstractNum*, num* — keep abstractNums grouped.
            var lastAbstract = root.Elements(W + "abstractNum").LastOrDefault();
            if (lastAbstract is not null) lastAbstract.AddAfterSelf(abstractNum);
            else
            {
                var firstNum = root.Elements(W + "num").FirstOrDefault();
                if (firstNum is not null) firstNum.AddBeforeSelf(abstractNum);
                else root.Add(abstractNum);
            }
        }

        var abstractId = (string)abstractNum.Attribute(W + "abstractNumId")!;

        // Reuse an existing w:num pointing at our abstractNum, or create one.
        var num = root.Elements(W + "num")
            .FirstOrDefault(n => (string?)n.Element(W + "abstractNumId")?.Attribute(W + "val") == abstractId);
        if (num is null)
        {
            int numId = NextId(root, "num", "numId");
            num = new XElement(W + "num",
                new XAttribute(W + "numId", numId),
                new XElement(W + "abstractNumId", new XAttribute(W + "val", abstractId)));
            root.Add(num); // nums come after abstractNums
        }

        // Flush the numbering part to its stream — the session's Save only persists the
        // projected parts (body/headers/...), not the numbering part we just mutated.
        part.PutXDocument();
        return (int)num.Attribute(W + "numId")!;
    }

    private static int NextId(XElement root, string elemLocalName, string idAttrLocalName)
    {
        int max = 0;
        foreach (var e in root.Elements(W + elemLocalName))
        {
            if (int.TryParse((string?)e.Attribute(W + idAttrLocalName), out var v))
                max = Math.Max(max, v);
        }
        return max + 1;
    }

    /// <summary>Build a spec-valid 9-level bullet or decimal abstractNum.</summary>
    private static XElement BuildAbstractNum(bool bullet, int absId, string nsid)
    {
        var an = new XElement(W + "abstractNum",
            new XAttribute(W + "abstractNumId", absId),
            new XElement(W + "nsid", new XAttribute(W + "val", nsid)),
            new XElement(W + "multiLevelType", new XAttribute(W + "val", "hybridMultilevel")));

        // Bullet glyphs cycle (•, o, ▪) using Symbol / Courier New / Wingdings, like Word.
        var bulletGlyphs = new[] { "", "o", "" };
        var bulletFonts = new[] { "Symbol", "Courier New", "Wingdings" };

        for (int lvl = 0; lvl < 9; lvl++)
        {
            int indentLeft = 720 * (lvl + 1);
            var pPr = new XElement(W + "pPr",
                new XElement(W + "ind",
                    new XAttribute(W + "left", indentLeft),
                    new XAttribute(W + "hanging", 360)));

            XElement lvl_;
            if (bullet)
            {
                lvl_ = new XElement(W + "lvl",
                    new XAttribute(W + "ilvl", lvl),
                    new XElement(W + "start", new XAttribute(W + "val", 1)),
                    new XElement(W + "numFmt", new XAttribute(W + "val", "bullet")),
                    new XElement(W + "lvlText", new XAttribute(W + "val", bulletGlyphs[lvl % 3])),
                    new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                    pPr,
                    new XElement(W + "rPr",
                        new XElement(W + "rFonts",
                            new XAttribute(W + "ascii", bulletFonts[lvl % 3]),
                            new XAttribute(W + "hAnsi", bulletFonts[lvl % 3]),
                            new XAttribute(W + "hint", "default"))));
            }
            else
            {
                lvl_ = new XElement(W + "lvl",
                    new XAttribute(W + "ilvl", lvl),
                    new XElement(W + "start", new XAttribute(W + "val", 1)),
                    new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")),
                    new XElement(W + "lvlText", new XAttribute(W + "val", $"%{lvl + 1}.")),
                    new XElement(W + "lvlJc", new XAttribute(W + "val", "left")),
                    pPr);
            }
            an.Add(lvl_);
        }

        return an;
    }
}
