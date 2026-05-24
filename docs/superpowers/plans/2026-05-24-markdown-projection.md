# Markdown Projection Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the `WmlToMarkdownConverter` scaffold with a full anchor-addressed Markdown projection of a DOCX, per `docs/architecture/markdown_projection.md`, propagated through the WASM bridge and npm wrapper.

**Architecture:** Single static `WmlToMarkdownConverter` class mirroring the *shape* (not bulk) of `WmlToHtmlConverter`. A first pass assigns Unids and builds an `AnchorIndex`; a second pass walks the document tree emitting Markdown, dispatching by element name to per-element handlers. A shared `UnidHelper` extracted from `WmlComparer` is used by both `WmlComparer` and the new converter so the Unid logic lives in one place. Multipart parts (body, headers, footers, footnotes, endnotes, comments) are emitted as named scope sections.

**Tech Stack:** .NET 8.0, DocumentFormat.OpenXml 3.2, xUnit, .NET WASM (`JSExport`), TypeScript, Playwright.

---

## File Structure

**Modify:**
- `Docxodus/WmlComparer.cs` — refactor `AssignUnidToAllElements` (line 8658) to call into a new shared helper.
- `Docxodus/WmlToMarkdownConverter.cs` — implement `Convert`, `AnchorTarget.Resolve`, and all phase logic. Becomes the bulk of the implementation (~1200–1800 lines projected).
- `docs/architecture/markdown_projection.md` — flip "Status: scaffold" → "Status: implemented (phases 1–8)" once finished.
- `CHANGELOG.md` — `[Unreleased]` entry.
- `wasm/DocxodusWasm/DocumentConverter.cs` — add `[JSExport]` markdown methods.
- `npm/src/types.ts` — add markdown types/enums + `DocxodusWasmExports` updates.
- `npm/src/index.ts` — add `convertWmlToMarkdown` and React-friendly wrapper.

**Create:**
- `Docxodus/UnidHelper.cs` — `UnidHelper.AssignToAllElements(XElement)` and `UnidHelper.GenerateUnid()`.
- `Docxodus.Tests/WmlToMarkdownConverterTests.cs` — all `MD###` tests.
- `npm/tests/markdown-projection.spec.ts` — Playwright tests for WASM/npm wrapper.

Tests reuse existing `TestFiles/HC*.docx` fixtures plus a few small programmatic fixtures.

---

## Task 1: Worktree + Branch + Compile Baseline

**Files:** none — environment only.

- [ ] **Step 1: Confirm branch.** `git status` must show `On branch feat/markdown-projection-impl`. (Branch was created by the orchestrator before this plan started.)
- [ ] **Step 2: Baseline build.** Run `dotnet build Docxodus.sln`. Must succeed (Debug build; warnings ok).
- [ ] **Step 3: Baseline tests.** Run `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName!~Wml" --logger "console;verbosity=normal"` quickly to confirm the test harness works in this checkout. (Skip Wml prefixed tests purely for speed.) Expect: existing tests pass.
- [ ] **Step 4: Confirm the existing scaffold actually throws.** `dotnet test --filter "FullyQualifiedName~WmlToMarkdown"` should return zero matched tests (no test file yet). That establishes our starting point.

---

## Task 2: Extract UnidHelper

**Files:**
- Create: `Docxodus/UnidHelper.cs`
- Modify: `Docxodus/WmlComparer.cs:8658-8684` (delete the private method body, call the helper instead)

- [ ] **Step 1: Write `Docxodus/UnidHelper.cs`:**

```csharp
#nullable enable

using System;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Shared helpers for the <c>PtOpenXml.Unid</c> stable-id attribute used by WmlComparer and
/// WmlToMarkdownConverter. The Unid is a 32-char hex string derived from a Guid; once assigned
/// to an element it remains stable across reformatting and reordering, which is the foundation
/// for the markdown projection's anchor scheme.
/// </summary>
internal static class UnidHelper
{
    /// <summary>Generate a fresh Unid value (32 hex chars).</summary>
    internal static string GenerateUnid() => Guid.NewGuid().ToString().Replace("-", "");

    /// <summary>
    /// Assign a <c>PtOpenXml.Unid</c> attribute to <paramref name="contentParent"/> (if it is a
    /// footnote/endnote root) and to every descendant that does not already have one.
    /// </summary>
    internal static void AssignToAllElements(XElement contentParent)
    {
        if (contentParent.Name == W.footnote || contentParent.Name == W.endnote)
        {
            if (contentParent.Attribute(PtOpenXml.Unid) == null)
            {
                contentParent.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
            }
        }

        foreach (var d in contentParent.Descendants())
        {
            if (d.Attribute(PtOpenXml.Unid) == null)
            {
                d.Add(new XAttribute(PtOpenXml.Unid, GenerateUnid()));
            }
        }
    }
}
```

- [ ] **Step 2: Replace `WmlComparer.AssignUnidToAllElements`.** In `Docxodus/WmlComparer.cs` at line 8658, replace the method body with a single call:

```csharp
private static void AssignUnidToAllElements(XElement contentParent)
{
    UnidHelper.AssignToAllElements(contentParent);
}
```

(Keep the wrapper to minimize churn at the ~6 call sites in WmlComparer; mark for future removal in a follow-up. Removing the wrapper here would require updating those call sites, which is out of scope.)

- [ ] **Step 3: Build.** `dotnet build Docxodus.sln` — expect success.
- [ ] **Step 4: Run existing WmlComparer tests for regression.** `dotnet test --filter "FullyQualifiedName~WmlComparerTests" -- RunConfiguration.MaxCpuCount=4` — must remain green.
- [ ] **Step 5: Commit.** `git add Docxodus/UnidHelper.cs Docxodus/WmlComparer.cs && git commit -m "refactor: extract UnidHelper from WmlComparer"`

---

## Task 3: Phase 1 — Anchor Index Skeleton

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Create: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Add tests file with first four cases.** New file `Docxodus.Tests/WmlToMarkdownConverterTests.cs`:

```csharp
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class WmlToMarkdownConverterTests
{
    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    [InlineData("HC004-ResumeTemplate.docx")]
    public void MD001_AnchorIndexIsExhaustive(string fixtureName)
    {
        var fi = new FileInfo(Path.Combine("../../../../TestFiles", fixtureName));
        var doc = new WmlDocument(fi.FullName);
        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart!.Document.Body!;
        var expectedBlocks = body.Descendants()
            .Where(d => d.LocalName == "p" || d.LocalName == "tbl")
            .Count();

        Assert.NotEmpty(projection.AnchorIndex);
        Assert.True(projection.AnchorIndex.Count >= expectedBlocks);
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD002_AnchorsAreStable(string fixtureName)
    {
        var fi = new FileInfo(Path.Combine("../../../../TestFiles", fixtureName));
        var doc1 = new WmlDocument(fi.FullName);
        var doc2 = new WmlDocument(fi.FullName);
        var p1 = WmlToMarkdownConverter.Convert(doc1, new WmlToMarkdownConverterSettings());
        var p2 = WmlToMarkdownConverter.Convert(doc2, new WmlToMarkdownConverterSettings());
        Assert.Equal(p1.AnchorIndex.Keys.OrderBy(k => k), p2.AnchorIndex.Keys.OrderBy(k => k));
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD003_AnchorsResolve(string fixtureName)
    {
        var fi = new FileInfo(Path.Combine("../../../../TestFiles", fixtureName));
        var doc = new WmlDocument(fi.FullName);
        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);

        foreach (var (id, target) in projection.AnchorIndex)
        {
            var element = target.Resolve(wdoc);
            Assert.NotNull(element);
            Assert.Equal(target.Unid, (string)element!.Attribute(PtOpenXml.Unid)!);
        }
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD004_AnchorsSurviveRoundTrip(string fixtureName)
    {
        var fi = new FileInfo(Path.Combine("../../../../TestFiles", fixtureName));
        var doc = new WmlDocument(fi.FullName);
        var first = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using (var stream = new MemoryStream())
        {
            stream.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
            using var wdoc = WordprocessingDocument.Open(stream, true);
            wdoc.MainDocumentPart!.Document.Save();
        }

        var second = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
        Assert.Equal(first.AnchorIndex.Keys.OrderBy(k => k), second.AnchorIndex.Keys.OrderBy(k => k));
    }
}
```

- [ ] **Step 2: Run and confirm failure.** `dotnet test --filter "FullyQualifiedName~WmlToMarkdownConverterTests"` — all four should fail with `NotImplementedException`.

- [ ] **Step 3: Implement Phase 1 in `Docxodus/WmlToMarkdownConverter.cs`.** Replace both `Convert` overloads' `NotImplementedException` bodies with anchor-index construction. Implementation outline:

```csharp
public static MarkdownProjection Convert(WmlDocument document, WmlToMarkdownConverterSettings settings)
{
    ArgumentNullException.ThrowIfNull(document);
    ArgumentNullException.ThrowIfNull(settings);

    using var stream = new MemoryStream();
    stream.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
    using var wdoc = WordprocessingDocument.Open(stream, true);
    return Convert(wdoc, settings);
}

public static MarkdownProjection Convert(WordprocessingDocument document, WmlToMarkdownConverterSettings settings)
{
    ArgumentNullException.ThrowIfNull(document);
    ArgumentNullException.ThrowIfNull(settings);

    // Phase 1: assign Unids everywhere and build the AnchorIndex.
    var index = BuildAnchorIndex(document, settings);
    // Phases 2+: emit markdown.
    var markdown = ""; // populated in later phases
    return new MarkdownProjection { Markdown = markdown, AnchorIndex = index };
}

private static IReadOnlyDictionary<string, AnchorTarget> BuildAnchorIndex(WordprocessingDocument doc, WmlToMarkdownConverterSettings settings)
{
    var index = new Dictionary<string, AnchorTarget>();
    var main = doc.MainDocumentPart ?? throw new InvalidOperationException("Document has no main part.");

    if (settings.Scopes.HasFlag(ProjectionScopes.Body))
        IndexPart(main, "body", main.GetXDocument().Root!, index);

    if (settings.Scopes.HasFlag(ProjectionScopes.Headers))
    {
        var i = 1;
        foreach (var hp in main.HeaderParts)
            IndexPart(hp, $"hdr{i++}", hp.GetXDocument().Root!, index);
    }
    if (settings.Scopes.HasFlag(ProjectionScopes.Footers))
    {
        var i = 1;
        foreach (var fp in main.FooterParts)
            IndexPart(fp, $"ftr{i++}", fp.GetXDocument().Root!, index);
    }
    if (settings.Scopes.HasFlag(ProjectionScopes.Footnotes) && main.FootnotesPart != null)
        IndexPart(main.FootnotesPart, "fn", main.FootnotesPart.GetXDocument().Root!, index);
    if (settings.Scopes.HasFlag(ProjectionScopes.Endnotes) && main.EndnotesPart != null)
        IndexPart(main.EndnotesPart, "en", main.EndnotesPart.GetXDocument().Root!, index);
    if (settings.Scopes.HasFlag(ProjectionScopes.Comments) && main.WordprocessingCommentsPart != null)
        IndexPart(main.WordprocessingCommentsPart, "cmt", main.WordprocessingCommentsPart.GetXDocument().Root!, index);

    // Persist Unid changes back to the package so the index can be resolved after save.
    main.PutXDocument();
    foreach (var hp in main.HeaderParts) hp.PutXDocument();
    foreach (var fp in main.FooterParts) fp.PutXDocument();
    main.FootnotesPart?.PutXDocument();
    main.EndnotesPart?.PutXDocument();
    main.WordprocessingCommentsPart?.PutXDocument();

    return index;
}

private static void IndexPart(OpenXmlPart part, string scope, XElement root, Dictionary<string, AnchorTarget> index)
{
    UnidHelper.AssignToAllElements(root);
    foreach (var el in root.DescendantsAndSelf())
    {
        var kind = KindFor(el);
        if (kind == null) continue;
        var unid = (string?)el.Attribute(PtOpenXml.Unid);
        if (unid == null) continue;
        var id = $"{kind}:{scope}:{unid}";
        var anchor = new Anchor(id, kind, scope, unid);
        index[id] = new AnchorTarget
        {
            Anchor = anchor,
            PartUri = part.Uri.ToString(),
            Unid = unid,
        };
    }
}

private static string? KindFor(XElement el)
{
    var n = el.Name;
    if (n == W.p) return IsHeading(el) ? "h" : (IsListItem(el) ? "li" : "p");
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
    return styleId != null && styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase);
}

private static bool IsListItem(XElement p) => p.Element(W.pPr)?.Element(W.numPr) != null;
```

- [ ] **Step 4: Implement `AnchorTarget.Resolve`.** Replace the `NotImplementedException`:

```csharp
public XElement? Resolve(WordprocessingDocument document)
{
    var uri = new Uri(PartUri, UriKind.Relative);
    var part = document.MainDocumentPart!.Parts
        .Select(p => p.OpenXmlPart)
        .Concat(document.MainDocumentPart.HeaderParts.Cast<OpenXmlPart>())
        .Concat(document.MainDocumentPart.FooterParts.Cast<OpenXmlPart>())
        .FirstOrDefault(p => p.Uri == uri)
        ?? (document.MainDocumentPart.Uri == uri ? document.MainDocumentPart : null);
    if (part == null) return null;
    var root = part.GetXDocument().Root;
    return root?.DescendantsAndSelf()
        .FirstOrDefault(e => (string?)e.Attribute(PtOpenXml.Unid) == Unid);
}
```

- [ ] **Step 5: Run Phase 1 tests.** `dotnet test --filter "FullyQualifiedName~WmlToMarkdownConverterTests"` — MD001–MD004 must pass.
- [ ] **Step 6: Commit.** `git commit -am "feat(markdown): phase 1 - anchor index + AnchorTarget.Resolve"`

---

## Task 4: Phase 2 — Paragraphs and Headings

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Modify: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Add tests `MD010`–`MD012`.** Add to `WmlToMarkdownConverterTests.cs`:

```csharp
[Fact]
public void MD010_BodyScopeHeaderEmitted()
{
    var doc = BuildSimpleDoc("Hello");
    var p = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.Body });
    Assert.StartsWith("# Document", p.Markdown);
}

[Fact]
public void MD011_ParagraphRendersWithAnchorOnOwnLine()
{
    var doc = BuildSimpleDoc("Hello world");
    var p = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
    Assert.Matches(@"\{#p:body:[0-9a-f]{32}\} Hello world", p.Markdown);
}

[Fact]
public void MD012_HeadingRendersAtCorrectLevel()
{
    var doc = BuildHeadingDoc("Title", 1);
    var p = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
    Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} # Title", p.Markdown);

    var p2 = WmlToMarkdownConverter.Convert(BuildHeadingDoc("Sub", 3),
        new WmlToMarkdownConverterSettings());
    Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} ### Sub", p2.Markdown);
}

[Fact]
public void MD013_HeadingLevelOffsetApplies()
{
    var doc = BuildHeadingDoc("Title", 1);
    var p = WmlToMarkdownConverter.Convert(doc,
        new WmlToMarkdownConverterSettings { HeadingLevelOffset = 1 });
    Assert.Matches(@"\{#h:body:[0-9a-f]{32}\} ## Title", p.Markdown);
}

[Fact]
public void MD014_AnchorModeNoneOmitsTokens()
{
    var doc = BuildSimpleDoc("Hello world");
    var p = WmlToMarkdownConverter.Convert(doc,
        new WmlToMarkdownConverterSettings { AnchorMode = AnchorRenderMode.None });
    Assert.DoesNotContain("{#p:", p.Markdown);
    Assert.Contains("Hello world", p.Markdown);
}
```

Add helpers `BuildSimpleDoc(string text)` and `BuildHeadingDoc(string text, int level)` that build minimal WmlDocuments using `DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create` writing into a `MemoryStream`, then return `new WmlDocument("fixture.docx", ms.ToArray())`. (Pattern lives in `DocumentBuilderTests.cs` — reuse there.) Heading paragraph needs a `w:pStyle w:val="Heading1"` and a corresponding style definition; add a minimal `Styles` part with one `w:style` per level you need.

- [ ] **Step 2: Run and confirm failure.** Tests should fail because Markdown is empty.

- [ ] **Step 3: Implement Phase 2 in `WmlToMarkdownConverter.cs`.** Add a Markdown emitter that walks each in-scope part and produces a StringBuilder. New helpers:

```csharp
private sealed class EmitContext
{
    public StringBuilder Sb { get; } = new();
    public WmlToMarkdownConverterSettings Settings { get; init; } = null!;
    public string Scope { get; set; } = "body";
}

private static string EmitMarkdown(WordprocessingDocument doc, WmlToMarkdownConverterSettings settings)
{
    var ctx = new EmitContext { Settings = settings };
    var main = doc.MainDocumentPart!;

    if (settings.Scopes.HasFlag(ProjectionScopes.Body))
    {
        ctx.Sb.AppendLine("# Document");
        ctx.Sb.AppendLine();
        ctx.Scope = "body";
        EmitBlocks(main.GetXDocument().Root!.Element(W.body)!.Elements(), ctx);
    }
    // ... headers, footers, footnotes, endnotes, comments (filled in Task 8)

    return ctx.Sb.ToString();
}

private static void EmitBlocks(IEnumerable<XElement> blocks, EmitContext ctx)
{
    foreach (var b in blocks)
    {
        if (b.Name == W.p) EmitParagraph(b, ctx);
        else if (b.Name == W.tbl) EmitTable(b, ctx); // Task 7
        else if (b.Name == W.sectPr) { /* Task 8 */ }
    }
}

private static void EmitParagraph(XElement p, EmitContext ctx)
{
    var anchor = ctx.Settings.AnchorMode == AnchorRenderMode.None ? "" : AnchorTokenFor(p, ctx) + " ";
    if (IsHeading(p))
    {
        var level = Math.Clamp(HeadingLevel(p) + ctx.Settings.HeadingLevelOffset, 1, 6);
        ctx.Sb.Append(anchor);
        ctx.Sb.Append(new string('#', level));
        ctx.Sb.Append(' ');
        EmitInlineRuns(p, ctx); // Task 5
        ctx.Sb.AppendLine();
        ctx.Sb.AppendLine();
        return;
    }
    if (IsListItem(p)) { EmitListItem(p, ctx); return; } // Task 6
    ctx.Sb.Append(anchor);
    EmitInlineRuns(p, ctx); // Task 5
    ctx.Sb.AppendLine();
    ctx.Sb.AppendLine();
}

private static string AnchorTokenFor(XElement el, EmitContext ctx)
{
    var kind = KindFor(el) ?? "unk";
    var unid = (string)el.Attribute(PtOpenXml.Unid)!;
    return $"{{#{kind}:{ctx.Scope}:{unid}}}";
}

private static int HeadingLevel(XElement p)
{
    var styleId = (string?)p.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val) ?? "Heading1";
    var digits = new string(styleId.Where(char.IsDigit).ToArray());
    return int.TryParse(digits, out var n) && n >= 1 && n <= 6 ? n : 1;
}
```

For Phase 2 only, `EmitInlineRuns` is the placeholder:

```csharp
private static void EmitInlineRuns(XElement p, EmitContext ctx)
{
    foreach (var r in p.Elements(W.r))
        foreach (var t in r.Elements(W.t))
            ctx.Sb.Append((string)t);
}
```

Replace the `markdown = ""` placeholder in `Convert(WordprocessingDocument, ...)` with `EmitMarkdown(document, settings)`.

- [ ] **Step 4: Run Phase 2 tests.** All MD010–MD014 must pass; MD001–MD004 must still pass.
- [ ] **Step 5: Commit.** `git commit -am "feat(markdown): phase 2 - paragraphs and headings"`

---

## Task 5: Phase 3 — Inline Runs

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Modify: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Tests `MD020`–`MD026`.** One per inline style. Build a helper `BuildRunsDoc(params (string text, bool bold, bool italic, bool code, bool strike)[] runs)`. Cases:

```csharp
[Fact] public void MD020_BoldRun() => AssertContains(BuildRunsDoc(("hello", true,false,false,false)), "**hello**");
[Fact] public void MD021_ItalicRun() => AssertContains(BuildRunsDoc(("hello", false,true,false,false)), "*hello*");
[Fact] public void MD022_CodeRun() => AssertContains(BuildRunsDoc(("x", false,false,true,false)), "`x`");
[Fact] public void MD023_StrikeRun() => AssertContains(BuildRunsDoc(("x", false,false,false,true)), "~~x~~");
[Fact] public void MD024_CombinedBoldItalic() => AssertContains(BuildRunsDoc(("hi", true,true,false,false)), "***hi***");
[Fact] public void MD025_BoldCancelsAcrossAdjacentRuns()
{
    var doc = BuildRunsDoc(("a", true, false, false, false), ("b", true, false, false, false));
    AssertContains(doc, "**ab**");          // merged, not "**a****b**"
}
[Fact] public void MD026_HyperlinkRendersAsLink()
{
    // Build doc with w:hyperlink referencing a relationship; assert "[text](https://example.com)"
}
[Fact] public void MD027_EscapesMarkdownMetacharacters()
{
    AssertContains(BuildRunsDoc(("a*b_c[d]", false,false,false,false)), @"a\*b\_c\[d\]");
}
```

`AssertContains(doc, expected)` runs the converter and asserts `Assert.Contains(expected, projection.Markdown)`.

- [ ] **Step 2: Run failures.**

- [ ] **Step 3: Implement inline emitter.** Replace the placeholder `EmitInlineRuns`:

```csharp
private static void EmitInlineRuns(XElement p, EmitContext ctx)
{
    // Group adjacent runs that share the same formatting so we emit one set of delimiters.
    var groups = GroupRunsByFormatting(p);
    foreach (var (fmt, runs) in groups)
    {
        var openClose = MarkdownDelimiters(fmt);
        if (fmt.IsHyperlink)
        {
            var url = ResolveHyperlinkUrl(fmt.HyperlinkRelId, fmt.HyperlinkAnchor, p);
            ctx.Sb.Append('[');
            foreach (var r in runs) AppendRunText(r, ctx);
            ctx.Sb.Append("](").Append(url).Append(')');
            continue;
        }
        ctx.Sb.Append(openClose.open);
        foreach (var r in runs) AppendRunText(r, ctx);
        ctx.Sb.Append(openClose.close);
    }
}

private record struct RunFormatting(bool Bold, bool Italic, bool Code, bool Strike, bool IsHyperlink, string? HyperlinkRelId, string? HyperlinkAnchor);

private static IEnumerable<(RunFormatting, List<XElement>)> GroupRunsByFormatting(XElement p) { /* … */ }
private static (string open, string close) MarkdownDelimiters(RunFormatting fmt) { /* combined delimiter logic, longer marker outside (e.g. *** ) */ }
private static void AppendRunText(XElement r, EmitContext ctx)
{
    foreach (var node in r.Elements())
    {
        if (node.Name == W.t) ctx.Sb.Append(EscapeMarkdown((string)node));
        else if (node.Name == W.br) ctx.Sb.Append("  \n"); // hard line break
        else if (node.Name == W.tab) ctx.Sb.Append("    ");
    }
}
private static string EscapeMarkdown(string s) =>
    System.Text.RegularExpressions.Regex.Replace(s, @"([\\`*_{}\[\]()#+\-.!|>~])", @"\$1");
```

Hyperlink URL resolution uses the part's relationships (`OpenXmlPart.HyperlinkRelationships`). For internal anchors (`w:hyperlink w:anchor=…`), emit `[text](#anchor)` — anchors here are Word bookmarks, not our Unid anchors.

- [ ] **Step 4: Pass tests, including MD001–MD026.**
- [ ] **Step 5: Commit.** `git commit -am "feat(markdown): phase 3 - inline runs (bold/italic/code/strike/links)"`

---

## Task 6: Phase 4 — Lists with Resolved Numbering

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Modify: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Tests `MD030`–`MD034`.** Reuse `TestFiles/HC001-5DayTourPlanTemplate.docx` and any HC fixture containing lists:

```csharp
[Theory]
[InlineData("HC001-5DayTourPlanTemplate.docx")]
public void MD030_BulletedListRendersWithDashMarkers(string fixture)
{
    var p = ConvertFixture(fixture);
    Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+-\s+\S", p.Markdown);
}

[Theory]
[InlineData("HC001-5DayTourPlanTemplate.docx")]
public void MD031_NumberedListRendersResolvedNumbers(string fixture)
{
    var p = ConvertFixture(fixture);
    Assert.Matches(@"\{#li:body:[0-9a-f]{32}\}\s+1\.\s", p.Markdown);
}

[Fact]
public void MD032_NestedListUsesIndentation()
{
    // small programmatic fixture: outer bullet, inner numPr with ilvl=1
    // expected: outer "- item" line, inner "  - item" indented 2 spaces
}

[Fact]
public void MD033_ResolveNumberingFalseFallsBackToDashes()
{
    var p = ConvertFixture("HC001-5DayTourPlanTemplate.docx", new WmlToMarkdownConverterSettings { ResolveNumbering = false });
    // No "1." literal markers; all list items use "-".
}

[Fact]
public void MD034_LegalNumberingRendersAsLiteralPrefix()
{
    // fixture with abstractNumId using "%1.%2." lvlText
    // expected line: "{#li:body:…} 1.2. item text"
}
```

- [ ] **Step 2: Implement list emitter.** Reuse `ListItemRetrieverSettings` infrastructure. Lazily compute list item information per part using existing `WmlToXml.AssembleListItemInformation` (line 1874). Cache the resolved level info on the `EmitContext` for reuse across paragraphs.

```csharp
private static void EmitListItem(XElement p, EmitContext ctx)
{
    var anchor = ctx.Settings.AnchorMode == AnchorRenderMode.None ? "" : AnchorTokenFor(p, ctx) + " ";
    var (ilvl, marker) = ResolveListMarker(p, ctx);
    var indent = new string(' ', ilvl * 2);
    ctx.Sb.Append(indent).Append(anchor).Append(marker).Append(' ');
    EmitInlineRuns(p, ctx);
    ctx.Sb.AppendLine();
    // Lists are NOT separated by blank lines (spec).
}

private static (int ilvl, string marker) ResolveListMarker(XElement p, EmitContext ctx)
{
    var ilvl = (int?)p.Element(W.pPr)?.Element(W.numPr)?.Element(W.ilvl)?.Attribute(W.val) ?? 0;
    if (!ctx.Settings.ResolveNumbering) return (ilvl, "-");
    // Resolve via ListItemRetriever / numbering.xml; fall back to "-" if format not recognized.
    var resolved = TryResolveLiteralListText(p, ctx);
    return (ilvl, resolved ?? "-");
}
```

Implementation note: full reuse of `ListItemRetriever` requires running `AssembleListItemInformation` on the `WordprocessingDocument` before walking. Do that once at the start of `EmitMarkdown` and stash the resulting list-item annotations on each paragraph (the existing implementation adds them as XAttributes named `ListItemRun.NumberAttributeName` or similar — inspect `WmlToXml.cs:1874` to confirm the exact attribute name).

If `ListItemRetriever` isn't ergonomic to reuse standalone, ship Phase 4 with the simpler behavior: bullets → `-`, ordered → `1.` / `a.` / `i.` based on the level's `w:lvlText`/`w:numFmt`, and document the limitation in `markdown_projection.md` under "Open Questions". Update the tests accordingly to a reduced assertion set, but do not skip MD030/MD031.

- [ ] **Step 3: Pass tests.**
- [ ] **Step 4: Commit.** `git commit -am "feat(markdown): phase 4 - lists with resolved numbering"`

---

## Task 7: Phase 5 — Tables (GFM + Opaque Fallback)

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Modify: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Tests `MD040`–`MD044`.**

```csharp
[Fact] public void MD040_SimpleTableRendersAsGfmPipeTable()
{
    // 2x2 fixture, plain text. Expect lines like:
    // | a | b |
    // | --- | --- |
    // | c | d |
}
[Fact] public void MD041_TableWithMergedCellsFallsBackToOpaque()
{
    // gridSpan in row 0 forces opaque block.
    // Expect: {#tbl:body:…}\n```table\nrows: 2\ncols: 2\n```
}
[Fact] public void MD042_NestedTableTriggersOpaqueFallback() { /* nested w:tbl inside w:tc */ }
[Fact] public void MD043_TableInlineCellMaxTriggersOpaque()
{
    // settings.TableInlineCellMax = 5; cell content of 6 chars → opaque
}
[Fact] public void MD044_TableModeAlwaysOpaqueAlwaysOpaque()
{
    // even simple 2x2 → opaque when TableMode = AlwaysOpaque
}
```

- [ ] **Step 2: Implement table emitter.**

```csharp
private static void EmitTable(XElement tbl, EmitContext ctx)
{
    var anchor = AnchorTokenFor(tbl, ctx);
    var mode = ctx.Settings.TableMode;
    if (mode == TableRenderMode.AlwaysOpaque || !CanRenderAsGfm(tbl, ctx))
    {
        EmitOpaqueTable(tbl, anchor, ctx);
        return;
    }
    EmitGfmTable(tbl, anchor, ctx);
}

private static bool CanRenderAsGfm(XElement tbl, EmitContext ctx)
{
    if (ctx.Settings.TableMode == TableRenderMode.AlwaysGfm) return true;
    // Disqualifiers: w:gridSpan with val>1, w:vMerge, nested w:tbl inside w:tc, cell content > TableInlineCellMax chars.
    return !tbl.Descendants(W.gridSpan).Any(g => (int?)g.Attribute(W.val) > 1)
        && !tbl.Descendants(W.vMerge).Any()
        && !tbl.Descendants(W.tc).Any(tc => tc.Descendants(W.tbl).Any())
        && tbl.Descendants(W.tc).All(tc => CellText(tc).Length <= ctx.Settings.TableInlineCellMax);
}

private static void EmitGfmTable(XElement tbl, string anchor, EmitContext ctx)
{
    ctx.Sb.AppendLine(anchor);
    var rows = tbl.Elements(W.tr).ToList();
    var first = rows.FirstOrDefault();
    if (first == null) return;
    var headerCells = first.Elements(W.tc).Select(tc => CellText(tc)).ToList();
    ctx.Sb.Append("| ").Append(string.Join(" | ", headerCells)).AppendLine(" |");
    ctx.Sb.Append("|").Append(string.Join("|", headerCells.Select(_ => " --- "))).AppendLine("|");
    foreach (var r in rows.Skip(1))
    {
        var cells = r.Elements(W.tc).Select(tc => CellText(tc));
        ctx.Sb.Append("| ").Append(string.Join(" | ", cells)).AppendLine(" |");
    }
    ctx.Sb.AppendLine();
}

private static void EmitOpaqueTable(XElement tbl, string anchor, EmitContext ctx)
{
    var rows = tbl.Elements(W.tr).Count();
    var cols = tbl.Elements(W.tr).FirstOrDefault()?.Elements(W.tc).Count() ?? 0;
    ctx.Sb.AppendLine(anchor);
    ctx.Sb.AppendLine("```table");
    ctx.Sb.Append("rows: ").Append(rows).AppendLine();
    ctx.Sb.Append("cols: ").Append(cols).AppendLine();
    ctx.Sb.AppendLine("```");
    ctx.Sb.AppendLine();
}

private static string CellText(XElement tc) =>
    string.Concat(tc.Descendants(W.t).Select(t => (string)t)).Replace("|", @"\|").Trim();
```

- [ ] **Step 3: Pass tests.**
- [ ] **Step 4: Commit.** `git commit -am "feat(markdown): phase 5 - tables (GFM + opaque fallback)"`

---

## Task 8: Phase 6 — Multipart Scopes

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Modify: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Tests `MD050`–`MD054`.**

```csharp
[Fact] public void MD050_HeadersSectionEmittedWhenPresent() { /* fixture with header */ }
[Fact] public void MD051_FootersSectionEmittedWhenPresent() { /* fixture with footer */ }
[Fact] public void MD052_FootnoteEmittedAsGfmFootnote() { /* "[^fn-...]" ref + definition list */ }
[Fact] public void MD053_EndnoteEmittedAsGfmFootnote() { /* "[^en-...]" */ }
[Fact] public void MD054_ScopesBodyOnlySkipsOtherSections()
{
    var p = ConvertFixture("HC001-...", new WmlToMarkdownConverterSettings { Scopes = ProjectionScopes.Body });
    Assert.DoesNotContain("# Headers", p.Markdown);
    Assert.DoesNotContain("# Footers", p.Markdown);
}
```

- [ ] **Step 2: Extend `EmitMarkdown`** to emit `# Headers / ## hdr{N}`, `# Footers / ## ftr{N}`, `# Footnotes`, `# Endnotes`, `# Comments` sections, each gated on the corresponding `ProjectionScopes` flag and the existence of the part.

Footnote references inside body paragraphs need to emit `[^fn-{unidShort}]` markers inline. Footnote definitions list under `# Footnotes` use the same anchor with `{#fn:fn:…}` prefix. Mirror for endnotes.

Comments: in body, wrap commented spans with inline `{#cmt:cmt:unid}` markers if `AnchorMode = BlockAndInline`. Under `# Comments`, list each as `- {#cmt:cmt:unid} **Author** (date): text`.

- [ ] **Step 3: Pass tests.**
- [ ] **Step 4: Commit.** `git commit -am "feat(markdown): phase 6 - multipart scopes (headers/footers/footnotes/endnotes/comments)"`

---

## Task 9: Phase 7 — Tracked Changes Modes

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs`
- Modify: `Docxodus.Tests/WmlToMarkdownConverterTests.cs`

- [ ] **Step 1: Tests `MD060`–`MD062`.**

```csharp
[Fact] public void MD060_AcceptModeDropsRevisionMarkup() { /* w:ins child contents present; w:del contents absent */ }
[Fact] public void MD061_RenderInlineModeEmitsBraceDelimiters()
{
    // Doc with w:ins "added" and w:del "gone" → "{+added+}", "{-gone-}"
}
[Fact] public void MD062_StripDeletionsKeepsInsertionsDropsDeletions() { /* w:ins kept, w:del dropped */ }
```

- [ ] **Step 2: Implement.** Before emitting, if `TrackedChanges == Accept`, run `RevisionAccepter.AcceptRevisions` on a clone (or pre-process the in-memory XDocument). For `RenderInline`, modify `AppendRunText` / paragraph walker to detect `w:ins` / `w:del` ancestors and wrap text accordingly. For `StripDeletions`, skip `w:del` descendants when walking.

- [ ] **Step 3: Pass tests.**
- [ ] **Step 4: Commit.** `git commit -am "feat(markdown): phase 7 - tracked changes modes"`

---

## Task 10: Phase 8 — WASM JSExport + npm wrapper

**Files:**
- Modify: `wasm/DocxodusWasm/DocumentConverter.cs`
- Modify: `npm/src/types.ts`
- Modify: `npm/src/index.ts`
- Create: `npm/tests/markdown-projection.spec.ts`

- [ ] **Step 1: Inspect existing WASM patterns.** Open `wasm/DocxodusWasm/DocumentConverter.cs` and `npm/src/types.ts`; locate the existing JSExport methods for `WmlToHtml` and `OpenContractExporter`. Mirror their shape.

- [ ] **Step 2: Add JSExport method.** In `wasm/DocxodusWasm/DocumentConverter.cs`:

```csharp
[JSExport]
public static string ConvertWmlToMarkdown(byte[] docxBytes, string settingsJson)
{
    var settings = string.IsNullOrWhiteSpace(settingsJson)
        ? new WmlToMarkdownConverterSettings()
        : JsonSerializer.Deserialize<MarkdownProjectionSettingsDto>(settingsJson)!.ToSettings();
    var doc = new WmlDocument("input.docx", docxBytes);
    var projection = WmlToMarkdownConverter.Convert(doc, settings);
    return JsonSerializer.Serialize(MarkdownProjectionDto.From(projection));
}
```

Add DTOs to flatten `AnchorIndex` for JSON: `MarkdownProjectionDto { string Markdown; Dictionary<string, AnchorTargetDto> AnchorIndex; }`.

- [ ] **Step 3: Add TypeScript types** in `npm/src/types.ts`:

```ts
export enum ProjectionScopes { Body = 1, Headers = 2, Footers = 4, Footnotes = 8, Endnotes = 16, Comments = 32, All = 63 }
export enum AnchorRenderMode { Block = 0, BlockAndInline = 1, None = 2 }
export enum TableRenderMode { GfmWithOpaqueFallback = 0, AlwaysGfm = 1, AlwaysOpaque = 2 }
export enum TrackedChangeMode { Accept = 0, RenderInline = 1, StripDeletions = 2 }

export interface WmlToMarkdownConverterSettings {
  scopes?: ProjectionScopes;
  headingLevelOffset?: number;
  anchorMode?: AnchorRenderMode;
  tableMode?: TableRenderMode;
  tableInlineCellMax?: number;
  trackedChanges?: TrackedChangeMode;
  resolveNumbering?: boolean;
}

export interface AnchorTarget { id: string; kind: string; scope: string; unid: string; partUri: string; }
export interface MarkdownProjection { markdown: string; anchorIndex: Record<string, AnchorTarget>; }
```

Extend `DocxodusWasmExports` with `ConvertWmlToMarkdown(bytes: Uint8Array, settingsJson: string): string`.

- [ ] **Step 4: Add wrapper** in `npm/src/index.ts`:

```ts
export async function convertWmlToMarkdown(docx: ArrayBuffer | Uint8Array, settings: WmlToMarkdownConverterSettings = {}): Promise<MarkdownProjection> {
  const exports = await getExports();
  const bytes = docx instanceof Uint8Array ? docx : new Uint8Array(docx);
  const json = exports.ConvertWmlToMarkdown(bytes, JSON.stringify(settings));
  return JSON.parse(json) as MarkdownProjection;
}
```

- [ ] **Step 5: Build npm.** `cd npm && npm run build`. Fix any compile errors.

- [ ] **Step 6: Playwright test.** Create `npm/tests/markdown-projection.spec.ts`:

```ts
import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
// follow the existing pattern in other npm/tests/*.spec.ts files for harness setup
test('convertWmlToMarkdown returns markdown and anchorIndex', async ({ page }) => {
  await page.goto('/');
  const fixturePath = resolve(__dirname, '../../TestFiles/HC001-5DayTourPlanTemplate.docx');
  const bytes = readFileSync(fixturePath);
  const result = await page.evaluate(async (b) => {
    const { convertWmlToMarkdown } = (window as any).Docxodus;
    return await convertWmlToMarkdown(new Uint8Array(b));
  }, Array.from(bytes));
  expect(result.markdown).toContain('# Document');
  expect(Object.keys(result.anchorIndex).length).toBeGreaterThan(0);
});
```

- [ ] **Step 7: Run.** `cd npm && npm test` should pass that spec alongside the existing ones.

- [ ] **Step 8: Commit.** `git commit -am "feat(markdown): phase 8 - WASM JSExport + npm wrapper"`

---

## Task 11: Documentation & Changelog

**Files:**
- Modify: `docs/architecture/markdown_projection.md`
- Modify: `CHANGELOG.md`
- Modify: `CLAUDE.md` (only if the public surface changed beyond what's already documented)

- [ ] **Step 1: Update spec status.** Change the top of `markdown_projection.md` from "Status: Design — scaffold only" to "Status: Implemented (phases 1–8). See history in CHANGELOG.md.". Strike the "Implementation Plan (Phases)" header and replace with "Implementation Phases — Done".

- [ ] **Step 2: Add `CHANGELOG.md` entry under `[Unreleased]`:**

```markdown
### Added
- `WmlToMarkdownConverter` now ships a full anchor-addressed Markdown projection of DOCX, replacing the v5.5.4 scaffold. Includes paragraphs, headings, inline runs, lists with resolved numbering, GFM tables with opaque fallback, multipart scopes (headers/footers/footnotes/endnotes/comments), tracked-change rendering modes, and a WASM/npm wrapper (`convertWmlToMarkdown`).
```

- [ ] **Step 3: Commit.** `git commit -am "docs(markdown): update spec status, changelog for markdown projection"`

---

## Task 12: Full Test Suite & PR

- [ ] **Step 1: Release build.** `dotnet build -c Release Docxodus.sln`. Must pass with TreatWarningsAsErrors.
- [ ] **Step 2: Full .NET test run.** `dotnet test Docxodus.Tests/Docxodus.Tests.csproj`. All previously-green tests stay green; all MD### tests pass.
- [ ] **Step 3: npm test run.** `cd npm && npm run build && npm test`. All Playwright specs pass.
- [ ] **Step 4: Push branch.** `git push -u origin feat/markdown-projection-impl`.
- [ ] **Step 5: PR.** `gh pr create --title "feat(markdown): full WmlToMarkdownConverter implementation" --body "Implements all 8 phases of docs/architecture/markdown_projection.md."`

---

## Notes

- **TDD discipline:** every phase ships test-first. Don't write Markdown emission before its test is failing.
- **Performance budget (spec §"Performance Budget"):** measure phase 1 cold time on `HC001` and the largest HC fixture. If >2× the targets, surface in the PR before adding more functionality.
- **Open question MD004 mitigation:** Task 3 step 3 already persists Unids back with `PutXDocument()`. Verify with MD004; if it still fails, investigate `PutXDocument` semantics inside `OpenXmlMemoryStreamDocument`.
- **`ListItemRetriever` fallback** (Task 6 step 2): if direct reuse is hard, the reduced behavior is acceptable for v1 and is called out in the spec's Open Questions.
