#nullable enable

// Regression tests for the bugs uncovered by the PR131 smoke test on the
// NVCA-Model-COI fixture. Each test pins one defect and (once the fix lands)
// becomes a guardrail. Numbering continues from DocxSessionTests in the
// "phase 10 — PR131 smoke fixes" range (DS080-DS099).

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionPR131RegressionTests
{
    // ─── Fixture builders ────────────────────────────────────────────────

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>Single body paragraph containing: text "AB" + hyperlink("CD") + text "EF".
    /// Full visible text is "ABCDEF" (6 chars).</summary>
    internal static byte[] BuildParagraphWithHyperlink()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var rel = main.AddHyperlinkRelationship(new System.Uri("https://example.com/cd"), true);

            var paraXml = new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "AB")),
                new XElement(W + "hyperlink",
                    new XAttribute(R + "id", rel.Id),
                    new XElement(W + "r",
                        new XElement(W + "t",
                            new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "CD"))),
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "EF")));

            var docXml = new XElement(W + "document",
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "w", W.NamespaceName),
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "r", R.NamespaceName),
                new XElement(W + "body", paraXml));
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    /// <summary>Single body paragraph wrapped by a bookmarkStart/End named "MARK1".
    /// Layout: bookmarkStart(id=1, name="MARK1") + run("hello world") + bookmarkEnd(id=1).</summary>
    internal static byte[] BuildParagraphWithBookmark()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var paraXml = new XElement(W + "p",
                new XElement(W + "bookmarkStart",
                    new XAttribute(W + "id", "1"),
                    new XAttribute(W + "name", "MARK1")),
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "hello world")),
                new XElement(W + "bookmarkEnd",
                    new XAttribute(W + "id", "1")));

            var docXml = new XElement(W + "document",
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "w", W.NamespaceName),
                new XElement(W + "body", paraXml));
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    /// <summary>Two body paragraphs: "First." and "Second has a hyperlink to [TXT]".
    /// Second paragraph contains: text "Second has a hyperlink to " + hyperlink("TXT").</summary>
    internal static byte[] BuildTwoParagraphsSecondWithHyperlink()
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            var rel = main.AddHyperlinkRelationship(new System.Uri("https://example.com/txt"), true);

            var p1 = new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "First.")));
            var p2 = new XElement(W + "p",
                new XElement(W + "r",
                    new XElement(W + "t",
                        new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "Second has a hyperlink to ")),
                new XElement(W + "hyperlink",
                    new XAttribute(R + "id", rel.Id),
                    new XElement(W + "r",
                        new XElement(W + "t",
                            new XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), "TXT"))));

            var docXml = new XElement(W + "document",
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "w", W.NamespaceName),
                new XAttribute(System.Xml.Linq.XNamespace.Xmlns + "r", R.NamespaceName),
                new XElement(W + "body", p1, p2));
            main.PutXDocument(new XDocument(docXml));
        }
        return ms.ToArray();
    }

    private static AnchorTarget FirstBodyParagraph(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind is "p" or "h" or "li" && t.Anchor.Scope == "body");

    private static System.Collections.Generic.List<AnchorTarget> BodyParagraphs(DocxSession s) =>
        s.Project().AnchorIndex.Values
            .Where(t => t.Anchor.Kind is "p" or "h" or "li" && t.Anchor.Scope == "body")
            .ToList();

    // ─── B4: bookmark preservation ───────────────────────────────────────

    [Fact]
    public void DS080_ReplaceText_PreservesBookmarks()
    {
        using var session = new DocxSession(BuildParagraphWithBookmark());
        var anchor = FirstBodyParagraph(session);

        var r = session.ReplaceText(anchor.Anchor.Id, "replaced text");
        Assert.True(r.Success, r.Error?.Message);

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var bookmarks = doc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W + "bookmarkStart")
            .Select(b => (string?)b.Attribute(W + "name"))
            .ToList();
        Assert.Contains("MARK1", bookmarks);
    }

    // ─── B1/B5: SplitParagraph hyperlink-aware ───────────────────────────

    [Fact]
    public void DS081_SplitParagraph_OffsetBetweenHyperlinkAndTrailingText()
    {
        // "AB" + hyperlink("CD") + "EF"  →  split at offset 4 (just after "CD")
        // Expected first = "ABCD" (hyperlink stays with first half),
        //          second = "EF".
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        Assert.Equal("ABCDEF", session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview);

        var r = session.SplitParagraph(anchor.Anchor.Id, 4);
        Assert.True(r.Success, $"split failed: {r.Error?.Code}/{r.Error?.Message}");

        var firstText = session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview;
        var secondText = session.GetAnchorInfo(r.Created[0].Id)?.TextPreview;
        Assert.Equal("ABCD", firstText);
        Assert.Equal("EF", secondText);
    }

    [Fact]
    public void DS082_SplitParagraph_OffsetAtPreviewLengthSucceeds()
    {
        // "AB" + hyperlink("CD") + "EF"  →  split at offset 6 (end of text)
        // ParagraphText (used to validate) MUST include hyperlink text,
        // otherwise the agent gets OffsetOutOfRange on a valid offset.
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        Assert.Equal(6, session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview.Length);

        var r = session.SplitParagraph(anchor.Anchor.Id, 6);
        Assert.True(r.Success, $"split at preview length must succeed; got {r.Error?.Code}/{r.Error?.Message}");
    }

    [Fact]
    public void DS083_SplitParagraph_OffsetInsideHyperlink()
    {
        // "AB" + hyperlink("CD") + "EF"  →  split at offset 3 (between C and D)
        // Expected: hyperlink content "CD" is split — first contains "ABC",
        // second contains "DEF".
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);

        var r = session.SplitParagraph(anchor.Anchor.Id, 3);
        Assert.True(r.Success, r.Error?.Message);

        var firstText = session.GetAnchorInfo(anchor.Anchor.Id)?.TextPreview;
        var secondText = session.GetAnchorInfo(r.Created[0].Id)?.TextPreview;
        Assert.Equal("ABC", firstText);
        Assert.Equal("DEF", secondText);
    }

    // ─── B2: MergeParagraphs preserves hyperlinks ────────────────────────

    [Fact]
    public void DS084_MergeParagraphs_PreservesHyperlinkInSecondParagraph()
    {
        using var session = new DocxSession(BuildTwoParagraphsSecondWithHyperlink());
        var paras = BodyParagraphs(session);
        Assert.True(paras.Count >= 2);
        var firstId = paras[0].Anchor.Id;
        var secondId = paras[1].Anchor.Id;

        var r = session.MergeParagraphs(firstId, secondId);
        Assert.True(r.Success, r.Error?.Message);

        var merged = session.GetAnchorInfo(firstId)?.TextPreview ?? "";
        Assert.Contains("TXT", merged);
    }

    // ─── B3: MergeParagraphs separator ───────────────────────────────────

    [Fact]
    public void DS085_MergeParagraphs_InsertsSeparator_WhenBothEndsAreNonWhitespace()
    {
        // Use the simple two-paragraph fixture ("First paragraph." + "Second paragraph.")
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var paras = BodyParagraphs(session);
        var firstId = paras[0].Anchor.Id;
        var secondId = paras[1].Anchor.Id;

        var r = session.MergeParagraphs(firstId, secondId);
        Assert.True(r.Success, r.Error?.Message);

        var merged = session.GetAnchorInfo(firstId)?.TextPreview;
        Assert.Equal("First paragraph. Second paragraph.", merged);
    }

    [Fact]
    public void DS086_MergeParagraphs_NoDoubleSpace_WhenFirstEndsWithWhitespace()
    {
        // Build a two-para fixture where the first ends with a trailing space.
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document();
            main.Document.Body = new Body();
            main.AddNewPart<StyleDefinitionsPart>().Styles = DocxSessionTests.BuildHeadingStyles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Body.Append(new Paragraph(new Run(new Text("First. ") { Space = SpaceProcessingModeValues.Preserve })));
            main.Document.Body.Append(new Paragraph(new Run(new Text("Second."))));
            main.Document.Save();
        }
        using var session = new DocxSession(ms.ToArray());
        var paras = BodyParagraphs(session);
        var r = session.MergeParagraphs(paras[0].Anchor.Id, paras[1].Anchor.Id);
        Assert.True(r.Success, r.Error?.Message);
        var merged = session.GetAnchorInfo(paras[0].Anchor.Id)?.TextPreview;
        Assert.Equal("First. Second.", merged);
    }

    // ─── B5: ApplyFormat hyperlink-aware ─────────────────────────────────

    [Fact]
    public void DS087_ApplyFormat_AcceptsSpanCoveringHyperlinkText()
    {
        // Preview is 6 chars ("ABCDEF"). ApplyFormat(0, 6) must succeed.
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        var r = session.ApplyFormat(anchor.Anchor.Id, new CharSpan(0, 6), new FormatOp { Bold = true });
        Assert.True(r.Success, $"got {r.Error?.Code}/{r.Error?.Message}");
    }

    [Fact]
    public void DS088_ApplyFormat_FormatsRunsInsideHyperlink()
    {
        using var session = new DocxSession(BuildParagraphWithHyperlink());
        var anchor = FirstBodyParagraph(session);
        var r = session.ApplyFormat(anchor.Anchor.Id, null, new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var hyperlinkRun = doc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W + "hyperlink").First()
            .Element(W + "r");
        var bold = hyperlinkRun?.Element(W + "rPr")?.Element(W + "b");
        Assert.NotNull(bold);
    }

    // ─── B6: hyperlink dedup ─────────────────────────────────────────────

    [Fact]
    public void DS089_ReplaceText_DedupesHyperlinkRelationship_SameUrl()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        for (int i = 0; i < 5; i++)
            session.ReplaceText(anchor.Anchor.Id, $"Round {i} text [link](https://example.com/same)");

        var saved = session.Save();
        using var doc = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var rels = doc.MainDocumentPart!.HyperlinkRelationships
            .Where(rl => rl.Uri.ToString() == "https://example.com/same")
            .ToList();
        Assert.Single(rels);
    }

    // ─── F1: bullet payload kind ─────────────────────────────────────────

    [Fact]
    public void DS090_InsertParagraph_BulletPayload_CreatedKindMatchesProjection()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        // Fixture has no numbering definitions and no preceding list, so the bullet
        // payload cannot inherit numPr. The returned kind MUST match what the
        // projector then reports for the created anchor — i.e. "p".
        var r = session.InsertParagraph(anchor.Anchor.Id, Position.After, "- bullet payload");
        Assert.True(r.Success, r.Error?.Message);

        var createdAnchor = r.Created[0];
        var projection = session.Project();
        var projectorTarget = projection.AnchorIndex.Values.FirstOrDefault(t => t.Unid == createdAnchor.Unid);
        Assert.NotNull(projectorTarget);
        Assert.Equal(projectorTarget!.Anchor.Kind, createdAnchor.Kind);
    }
}
