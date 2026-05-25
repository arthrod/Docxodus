# Issue #162 — Anchor Introspection Ergonomics

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make anchor introspection ergonomic for agents working with `DocxSession` — surface text previews directly on `AnchorTarget`, exclude Word-reserved footnote separators from the projection's `AnchorIndex`, and add a bulk `GetAnchorInfos` lookup. Eliminates the per-anchor `GetAnchorInfo` round-trip pattern that costs N walks for a 500-anchor doc.

**Architecture:** Compute the text preview once during projection (in `BuildAnchorIndex`) and stash it on `AnchorTarget`. Pre-collect the set of boilerplate fn/en elements (and their descendants) before the descendant walk so they never become anchors. Refactor `DocxSession.GetAnchorInfo` to read `target.TextPreview` instead of re-walking the element. Add `GetAnchorInfos(IEnumerable<string>)` as a thin batched wrapper. Propagate `textPreview` through the WASM `MarkdownAnchorTargetDto` + bridge JSON serializers and the npm `MarkdownAnchorTarget` / `AnchorTargetRef` type shapes.

**Tech Stack:** .NET 8, C#, Open XML SDK 3.x, xUnit. WASM bridge `[JSExport]`, npm TypeScript wrapper. Playwright for browser tests.

**Branch:** `feat/162-anchor-introspection`

---

## Task 1: Create feature branch and confirm clean working tree

**Files:** none (git state only)

- [ ] **Step 1.1: Confirm clean tree on main**

```bash
cd /home/jman/Code/Docxodus
git status
git log --oneline -3
```

Expected: working tree clean (no uncommitted changes), on `main`.

- [ ] **Step 1.2: Create and switch to feature branch**

```bash
git checkout -b feat/162-anchor-introspection
```

Expected: `Switched to a new branch 'feat/162-anchor-introspection'`.

---

## Task 2: Add `TextPreview` to `AnchorTarget` (failing test)

**Files:**
- Test: `Docxodus.Tests/WmlToMarkdownConverterTests.cs` — append a new test method `MD040_AnchorTargetCarriesTextPreview`.

- [ ] **Step 2.1: Add the failing test**

Open `Docxodus.Tests/WmlToMarkdownConverterTests.cs`. Find the existing class and append this test method (use the existing `BuildDocx` helper pattern in that file — search for an existing `MD0` test to see the helper signature):

```csharp
[Fact]
public void MD040_AnchorTargetCarriesTextPreview()
{
    // A two-paragraph body — each AnchorTarget should expose the first ~80 chars
    // of the element's flat text directly, without needing a separate session walk.
    var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
    var wml = new WmlDocument("test.docx", bytes);

    var projection = WmlToMarkdownConverter.Convert(wml, new WmlToMarkdownConverterSettings());

    var bodyParas = projection.AnchorIndex.Values
        .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
        .ToList();

    Assert.NotEmpty(bodyParas);
    foreach (var t in bodyParas)
    {
        // Every body block has a non-null preview; for non-empty paragraphs it's non-empty.
        Assert.NotNull(t.TextPreview);
    }

    // At least one paragraph has the expected literal first chars
    // (BuildDS001 paragraphs start with "First paragraph" and "Second paragraph").
    Assert.Contains(bodyParas, t => t.TextPreview.StartsWith("First paragraph"));
    Assert.Contains(bodyParas, t => t.TextPreview.StartsWith("Second paragraph"));
}
```

- [ ] **Step 2.2: Run the test, verify it fails**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD040_AnchorTargetCarriesTextPreview" 2>&1 | tail -20
```

Expected: build fails with `'AnchorTarget' does not contain a definition for 'TextPreview'`.

---

## Task 3: Add `TextPreview` to `AnchorTarget` (implementation)

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs` lines 151–190 (the `AnchorTarget` class) and lines 302–337 (the `BuildAnchorIndex` walker).

- [ ] **Step 3.1: Add the `TextPreview` field to `AnchorTarget`**

In `Docxodus/WmlToMarkdownConverter.cs`, find the `AnchorTarget` class declaration around line 151. Add a new property after `Unid`:

```csharp
public sealed class AnchorTarget
{
    /// <summary>The anchor this target resolves.</summary>
    required public Anchor Anchor { get; init; }

    /// <summary>URI of the package part containing the element (e.g. main document, header part).</summary>
    required public string PartUri { get; init; }

    /// <summary>Stable Unid of the element. Use with Docxodus' Unid lookup to fetch the live XElement.</summary>
    required public string Unid { get; init; }

    /// <summary>
    /// First ~80 characters of the element's flat text, suitable for showing in
    /// agent context windows or UI lists. Computed during projection so agents
    /// don't need to <see cref="Resolve"/> + re-walk the element for previews.
    /// Empty for elements with no text (e.g. empty paragraphs, section breaks).
    /// </summary>
    public string TextPreview { get; init; } = string.Empty;

    /// <summary> ... existing Resolve method ... </summary>
    public XElement? Resolve(WordprocessingDocument document)
    // ... unchanged ...
}
```

- [ ] **Step 3.2: Add a private text-preview helper alongside `BuildAnchorIndex`**

In `Docxodus/WmlToMarkdownConverter.cs`, just before the `BuildAnchorIndex` method (around line 274), add this private static helper:

```csharp
private const int TextPreviewMaxLength = 80;

private static string ComputeTextPreview(XElement element)
{
    var text = string.Concat(element.Descendants(W.t).Select(t => (string)t));
    return text.Length > TextPreviewMaxLength
        ? text.Substring(0, TextPreviewMaxLength) + "…"   // "…"
        : text;
}
```

- [ ] **Step 3.3: Populate `TextPreview` when constructing each `AnchorTarget`**

In `BuildAnchorIndex` around line 326, change the `AnchorTarget` constructor call:

```csharp
// Before:
var anchor = new Anchor(id, kind, scope.Name, unid);
index[id] = new AnchorTarget
{
    Anchor = anchor,
    PartUri = scope.Part.Uri.ToString(),
    Unid = unid,
};

// After:
var anchor = new Anchor(id, kind, scope.Name, unid);
index[id] = new AnchorTarget
{
    Anchor = anchor,
    PartUri = scope.Part.Uri.ToString(),
    Unid = unid,
    TextPreview = ComputeTextPreview(el),
};
```

- [ ] **Step 3.4: Run the test, verify it passes**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD040_AnchorTargetCarriesTextPreview" 2>&1 | tail -10
```

Expected: `Passed!  - Failed: 0, Passed: 1, Skipped: 0`.

- [ ] **Step 3.5: Run the full markdown-projection test class to ensure no regressions**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~WmlToMarkdownConverterTests" 2>&1 | tail -5
```

Expected: all green.

- [ ] **Step 3.6: Commit**

```bash
git add Docxodus/WmlToMarkdownConverter.cs Docxodus.Tests/WmlToMarkdownConverterTests.cs
git commit -m "feat(markdown-projection): expose TextPreview on AnchorTarget (#162)"
```

---

## Task 4: Filter Word-reserved footnote separators from `AnchorIndex` (failing test)

**Files:**
- Test: `Docxodus.Tests/WmlToMarkdownConverterTests.cs` — append `MD041_BoilerplateFootnotesNotInAnchorIndex`.

- [ ] **Step 4.1: Add the failing test**

Append to `Docxodus.Tests/WmlToMarkdownConverterTests.cs`. We need a doc with footnotes; use the existing `DocxSessionTests` helpers if one exists, or build inline. Search the file first for a footnote-fixture helper:

```bash
grep -nE "footnote|FootnotesPart|w:footnote" /home/jman/Code/Docxodus/Docxodus.Tests/WmlToMarkdownConverterTests.cs /home/jman/Code/Docxodus/Docxodus.Tests/DocxSessionTests.cs | head -10
```

If a helper like `BuildDocWithFootnotes` exists, use it. Otherwise, the test can use the NVCA template if it's in `TestFiles/`:

```bash
ls Docxodus.Tests/TestFiles/ | grep -iE "nvca|footnote" | head -5
```

Append this test, adapting the fixture-load to whatever helper exists in the codebase (the assertion is the important part — the fixture mechanic varies by what's already available):

```csharp
[Fact]
public void MD041_BoilerplateFootnotesNotInAnchorIndex()
{
    // Every DOCX with a FootnotesPart includes two Word-reserved boilerplate
    // footnotes (type="separator" and type="continuationSeparator") used to
    // render the horizontal separator lines above footnote text. They are
    // structural plumbing — not editorial content — and must not appear in
    // the agent-facing AnchorIndex.
    //
    // Use any existing fixture that contains footnotes. Falls back to
    // programmatic construction if none exists.
    var bytes = DocxSessionTests.BuildDocWithFootnotes();   // if helper exists; otherwise see below
    var wml = new WmlDocument("test.docx", bytes);

    var projection = WmlToMarkdownConverter.Convert(wml, new WmlToMarkdownConverterSettings());

    // Resolve every fn-kind anchor in the fn scope back to its XElement
    // and assert none is a separator/continuationSeparator.
    using var sm = new OpenXmlMemoryStreamDocument(wml);
    using var doc = sm.GetWordprocessingDocument();
    foreach (var t in projection.AnchorIndex.Values
                 .Where(t => t.Anchor.Scope == "fn" && t.Anchor.Kind == "fn"))
    {
        var el = t.Resolve(doc);
        Assert.NotNull(el);
        var type = (string?)el.Attribute(W.type);
        Assert.False(type is "separator" or "continuationSeparator",
            $"Boilerplate fn (type={type}) leaked into AnchorIndex: {t.Anchor.Id}");
    }

    // And descendant paragraphs of boilerplate notes shouldn't appear either —
    // their scope is "fn" but their kind is "p"/"h"/"li". An empty set is also
    // valid (a doc with no real footnotes), so we just verify no leakage.
    foreach (var t in projection.AnchorIndex.Values
                 .Where(t => t.Anchor.Scope == "fn" && t.Anchor.Kind is "p" or "h" or "li"))
    {
        var el = t.Resolve(doc);
        Assert.NotNull(el);
        var ancestorFn = el.Ancestors(W.footnote).FirstOrDefault();
        Assert.NotNull(ancestorFn);
        var type = (string?)ancestorFn.Attribute(W.type);
        Assert.False(type is "separator" or "continuationSeparator",
            $"Descendant of boilerplate fn leaked into AnchorIndex: {t.Anchor.Id}");
    }
}
```

If `BuildDocWithFootnotes` does not exist, add this helper to `DocxSessionTests.cs` (or wherever the other `BuildDS00x` helpers live). It produces a minimal DOCX with one user footnote (Word auto-adds the boilerplate separators when a FootnotesPart is created):

```csharp
internal static byte[] BuildDocWithFootnotes()
{
    using var ms = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
            new DocumentFormat.OpenXml.Wordprocessing.Body(
                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(
                        new DocumentFormat.OpenXml.Wordprocessing.Text("Body text with a footnote reference."),
                        new DocumentFormat.OpenXml.Wordprocessing.FootnoteReference { Id = 1 }))));

        var fnPart = main.AddNewPart<DocumentFormat.OpenXml.Packaging.FootnotesPart>();
        var fnXml = """
            <w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
              <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
              <w:footnote w:id="1"><w:p><w:r><w:t>User-authored footnote text.</w:t></w:r></w:p></w:footnote>
            </w:footnotes>
            """;
        using (var s = fnPart.GetStream(FileMode.Create))
        using (var w = new StreamWriter(s)) w.Write(fnXml);
    }
    return ms.ToArray();
}
```

- [ ] **Step 4.2: Run the test, verify it fails**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD041_BoilerplateFootnotesNotInAnchorIndex" 2>&1 | tail -15
```

Expected: FAIL — current code anchors boilerplate footnotes and their paragraphs.

---

## Task 5: Filter Word-reserved footnote separators from `AnchorIndex` (implementation)

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs` — `BuildAnchorIndex` (around lines 302–337).

- [ ] **Step 5.1: Compute the skip-set before the descendant walk**

In `BuildAnchorIndex`, just inside the `foreach (var scope in scopes)` loop, before the `UnidHelper.AssignToAllElements(scope.Root);` line is the right place to think about this — but the existing call assigns Unids first. Add the skip-set computation **after** the existing root-annotation block (around line 309), **before** the `foreach (var el in scope.Root.DescendantsAndSelf())` loop.

Replace this block:

```csharp
foreach (var el in scope.Root.DescendantsAndSelf())
{
    var kind = KindFor(el);
    if (kind == null) continue;
    // ... existing body ...
}
```

With:

```csharp
// Word-reserved footnote/endnote separators (type="separator" / type="continuationSeparator")
// are structural plumbing that cannot be deleted and should not appear in the agent-facing
// AnchorIndex. Pre-collect them and their descendants so the walker skips both the notes
// themselves and any paragraphs/runs they contain.
var skip = new HashSet<XElement>();
if (scope.Name is "fn" or "en")
{
    var noteName = scope.Name == "fn" ? W.footnote : W.endnote;
    foreach (var n in scope.Root.Elements(noteName))
    {
        if (IsBoilerplateNote(n))
        {
            skip.Add(n);
            foreach (var d in n.Descendants()) skip.Add(d);
        }
    }
}

foreach (var el in scope.Root.DescendantsAndSelf())
{
    if (skip.Contains(el)) continue;
    var kind = KindFor(el);
    if (kind == null) continue;
    // ... existing body unchanged ...
}
```

`IsBoilerplateNote` already exists at WmlToMarkdownConverter.cs:528.

- [ ] **Step 5.2: Run the test, verify it passes**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD041_BoilerplateFootnotesNotInAnchorIndex" 2>&1 | tail -10
```

Expected: `Passed!`.

- [ ] **Step 5.3: Run the full WmlToMarkdownConverter + DocxSession test classes**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~WmlToMarkdownConverterTests|FullyQualifiedName~DocxSessionTests" 2>&1 | tail -10
```

Expected: all green. If any DS test breaks because it was counting boilerplate fn anchors, that's a test-side bug (counting plumbing as content) — fix the assertion, not the implementation.

- [ ] **Step 5.4: Commit**

```bash
git add Docxodus/WmlToMarkdownConverter.cs Docxodus.Tests/WmlToMarkdownConverterTests.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(markdown-projection): exclude Word-reserved footnote separators from AnchorIndex (#162)"
```

---

## Task 6: Add `GetAnchorInfos` bulk lookup and simplify `GetAnchorInfo` (failing test)

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS220_GetAnchorInfosBulkLookup`.

- [ ] **Step 6.1: Add the failing test**

Append to `Docxodus.Tests/DocxSessionTests.cs` (use the test-number range `DS220+` since the latest used was around DS207 per CHANGELOG):

```csharp
[Fact]
public void DS220_GetAnchorInfosBulkLookup()
{
    using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
    var projection = session.Project();

    // Collect every body paragraph anchor id.
    var ids = projection.AnchorIndex.Values
        .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li")
        .Select(t => t.Anchor.Id)
        .ToList();
    Assert.NotEmpty(ids);

    // Bulk lookup returns the same data as N individual GetAnchorInfo calls,
    // and an unknown anchor id maps to null in the result.
    var idsWithBogus = ids.Concat(new[] { "p:body:0000000000000000ffffffffffffffff" }).ToList();
    var bulk = session.GetAnchorInfos(idsWithBogus);

    Assert.Equal(idsWithBogus.Count, bulk.Count);
    foreach (var id in ids)
    {
        Assert.True(bulk.ContainsKey(id));
        var info = bulk[id];
        Assert.NotNull(info);
        Assert.Equal(id, info!.Id);

        // Verify it agrees with the existing per-anchor API.
        var singleton = session.GetAnchorInfo(id);
        Assert.NotNull(singleton);
        Assert.Equal(singleton!.TextPreview, info.TextPreview);
        Assert.Equal(singleton.Kind, info.Kind);
        Assert.Equal(singleton.Scope, info.Scope);
    }
    Assert.Null(bulk["p:body:0000000000000000ffffffffffffffff"]);
}

[Fact]
public void DS221_GetAnchorInfoUsesAnchorTargetTextPreview()
{
    // Regression: after #162, GetAnchorInfo reads target.TextPreview directly
    // instead of re-walking the element. Confirm the value matches what the
    // projection put on AnchorTarget.
    using var session = new DocxSession(BuildDS001_SimpleTwoParagraphs());
    var projection = session.Project();
    foreach (var t in projection.AnchorIndex.Values
                  .Where(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h" or "li"))
    {
        var info = session.GetAnchorInfo(t.Anchor.Id);
        Assert.NotNull(info);
        Assert.Equal(t.TextPreview, info!.TextPreview);
    }
}
```

- [ ] **Step 6.2: Run the new tests, verify they fail**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS220_GetAnchorInfosBulkLookup|FullyQualifiedName~DS221_GetAnchorInfoUsesAnchorTargetTextPreview" 2>&1 | tail -15
```

Expected: `DS220` build-fails with `'DocxSession' does not contain a definition for 'GetAnchorInfos'`; `DS221` may pass already (the value should agree since both read the same flat text) — that's fine.

---

## Task 7: Add `GetAnchorInfos` bulk lookup and simplify `GetAnchorInfo` (implementation)

**Files:**
- Modify: `Docxodus/DocxSession.cs` — `GetAnchorInfo` at line 378, plus add `GetAnchorInfos`. Remove `ElementTextPreview` at line 2257.

- [ ] **Step 7.1: Simplify `GetAnchorInfo` to use `AnchorTarget.TextPreview`**

In `Docxodus/DocxSession.cs`, replace the body of `GetAnchorInfo` (around line 378):

```csharp
// Before:
public AnchorInfo? GetAnchorInfo(string anchorId)
{
    ThrowIfDisposed();
    var target = FindAnchor(anchorId);
    if (target is null) return null;

    var element = target.Resolve(_doc!);
    var preview = element is null ? "" : ElementTextPreview(element);
    return new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, preview);
}

// After:
public AnchorInfo? GetAnchorInfo(string anchorId)
{
    ThrowIfDisposed();
    var target = FindAnchor(anchorId);
    if (target is null) return null;
    return new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, target.TextPreview);
}
```

- [ ] **Step 7.2: Add `GetAnchorInfos` directly below `GetAnchorInfo`**

```csharp
/// <summary>
/// Bulk variant of <see cref="GetAnchorInfo"/>. Resolves every requested anchor
/// from the projection's cached <c>AnchorIndex</c> in a single pass. Unknown
/// anchor ids map to <c>null</c> in the returned dictionary so callers can
/// distinguish "anchor doesn't exist" from "anchor exists with empty preview."
/// </summary>
public IReadOnlyDictionary<string, AnchorInfo?> GetAnchorInfos(IEnumerable<string> anchorIds)
{
    ThrowIfDisposed();
    ArgumentNullException.ThrowIfNull(anchorIds);

    var result = new Dictionary<string, AnchorInfo?>(StringComparer.Ordinal);
    foreach (var id in anchorIds)
    {
        if (id is null) continue;
        if (result.ContainsKey(id)) continue;
        var target = FindAnchor(id);
        result[id] = target is null
            ? null
            : new AnchorInfo(target.Anchor.Id, target.Anchor.Kind, target.Anchor.Scope, target.TextPreview);
    }
    return result;
}
```

- [ ] **Step 7.3: Remove the now-unused `ElementTextPreview` private helper**

At `Docxodus/DocxSession.cs:2257`, delete the entire method:

```csharp
private static string ElementTextPreview(XElement element)
{
    var text = string.Concat(element.Descendants(W.t).Select(t => (string)t));
    return text.Length > 80 ? text.Substring(0, 80) + "…" : text;
}
```

Confirm nothing else calls it first:

```bash
grep -n "ElementTextPreview" /home/jman/Code/Docxodus/Docxodus/*.cs /home/jman/Code/Docxodus/Docxodus.Tests/*.cs
```

Expected: no results after the deletion. If anything else references it, fix that caller to use `target.TextPreview` via `FindAnchor` instead.

- [ ] **Step 7.4: Run the tests, verify they pass**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS220|FullyQualifiedName~DS221|FullyQualifiedName~DS003" 2>&1 | tail -10
```

Expected: all three green (DS003 is the existing `ExistsAndGetAnchorInfo` regression).

- [ ] **Step 7.5: Run the entire DocxSession test class**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionTests" 2>&1 | tail -5
```

Expected: all green.

- [ ] **Step 7.6: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): add GetAnchorInfos bulk lookup, simplify GetAnchorInfo (#162)"
```

---

## Task 8: WASM bridge — propagate `TextPreview` through DTOs

**Files:**
- Modify: `wasm/DocxodusWasm/JsonContext.cs:1243-1250` — add `TextPreview` to `MarkdownAnchorTargetDto`.
- Modify: `wasm/DocxodusWasm/DocumentConverter.cs:898-907` — populate `TextPreview`.
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs:644-652` — emit `textPreview` in `AppendAnchorTarget`.

- [ ] **Step 8.1: Add `TextPreview` to the DTO**

In `wasm/DocxodusWasm/JsonContext.cs`, find `MarkdownAnchorTargetDto` (around line 1243) and add the field:

```csharp
public class MarkdownAnchorTargetDto
{
    public string Id { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public string Scope { get; set; } = string.Empty;
    public string Unid { get; set; } = string.Empty;
    public string PartUri { get; set; } = string.Empty;
    public string TextPreview { get; set; } = string.Empty;   // NEW
}
```

- [ ] **Step 8.2: Populate `TextPreview` in `DocumentConverter.cs`**

In `wasm/DocxodusWasm/DocumentConverter.cs` around line 898, update the projection DTO build:

```csharp
response.AnchorIndex[id] = new MarkdownAnchorTargetDto
{
    Id = target.Anchor.Id,
    Kind = target.Anchor.Kind,
    Scope = target.Anchor.Scope,
    Unid = target.Anchor.Unid,
    PartUri = target.PartUri,
    TextPreview = target.TextPreview,   // NEW
};
```

- [ ] **Step 8.3: Add `textPreview` to `AppendAnchorTarget` JSON serializer**

In `wasm/DocxodusWasm/DocxSessionBridge.cs` around line 644, update:

```csharp
private static void AppendAnchorTarget(StringBuilder sb, AnchorTarget t)
{
    sb.Append("{\"id\":").Append(JsonString(t.Anchor.Id))
      .Append(",\"kind\":").Append(JsonString(t.Anchor.Kind))
      .Append(",\"scope\":").Append(JsonString(t.Anchor.Scope))
      .Append(",\"unid\":").Append(JsonString(t.Unid))
      .Append(",\"partUri\":").Append(JsonString(t.PartUri))
      .Append(",\"textPreview\":").Append(JsonString(t.TextPreview))   // NEW
      .Append('}');
}
```

- [ ] **Step 8.4: Build the WASM target**

```bash
./scripts/build-wasm.sh 2>&1 | tail -15
```

Expected: builds cleanly. Note: after this runs, the .NET `Docxodus.dll` in `bin/Debug` is the WASM-mode build (no SkiaSharp). Tests will need a `dotnet clean` before going back to non-WASM workflow — see CLAUDE.md.

---

## Task 9: WASM bridge — add `GetAnchorInfo` / `GetAnchorInfos` JSExports

**Files:**
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs` — add two new `[JSExport]` methods.

- [ ] **Step 9.1: Add the JSExport methods**

In `wasm/DocxodusWasm/DocxSessionBridge.cs`, find a logical spot near the existing anchor-discovery methods (search for `FindByBookmark` around line 243 and insert below the block of related methods). Add:

```csharp
/// <summary>
/// Returns AnchorInfo as JSON for one anchor id, or "null" if the anchor
/// is unknown. Lets npm callers avoid re-projecting just to read a preview.
/// </summary>
[JSExport]
public static string GetAnchorInfo(int h, string anchorId)
{
    var info = Get(h).GetAnchorInfo(anchorId);
    if (info is null) return "null";
    var sb = new StringBuilder(200);
    sb.Append("{\"id\":").Append(JsonString(info.Id))
      .Append(",\"kind\":").Append(JsonString(info.Kind))
      .Append(",\"scope\":").Append(JsonString(info.Scope))
      .Append(",\"textPreview\":").Append(JsonString(info.TextPreview))
      .Append('}');
    return sb.ToString();
}

/// <summary>
/// Bulk variant: takes a JSON array of anchor ids and returns a JSON object
/// keyed by id (value is the same AnchorInfo shape, or null for unknown ids).
/// </summary>
[JSExport]
public static string GetAnchorInfos(int h, string anchorIdsJson)
{
    string[] ids;
    try
    {
        ids = JsonSerializer.Deserialize<string[]>(
            anchorIdsJson, DocxodusJsonContext.Default.StringArray) ?? Array.Empty<string>();
    }
    catch (JsonException)
    {
        return "{\"error\":\"malformed anchor id array\"}";
    }

    var map = Get(h).GetAnchorInfos(ids);
    var sb = new StringBuilder(ids.Length * 100 + 2);
    sb.Append('{');
    bool first = true;
    foreach (var kv in map)
    {
        if (!first) sb.Append(',');
        first = false;
        sb.Append(JsonString(kv.Key)).Append(':');
        if (kv.Value is null) sb.Append("null");
        else
        {
            sb.Append("{\"id\":").Append(JsonString(kv.Value.Id))
              .Append(",\"kind\":").Append(JsonString(kv.Value.Kind))
              .Append(",\"scope\":").Append(JsonString(kv.Value.Scope))
              .Append(",\"textPreview\":").Append(JsonString(kv.Value.TextPreview))
              .Append('}');
        }
    }
    sb.Append('}');
    return sb.ToString();
}
```

- [ ] **Step 9.2: Add `StringArray` to `DocxodusJsonContext` if not already present**

Check whether the source generator already knows about `string[]`:

```bash
grep -n "JsonSerializable(typeof(string\[\]))" /home/jman/Code/Docxodus/wasm/DocxodusWasm/JsonContext.cs
```

If not present, add to `JsonContext.cs`:

```csharp
[JsonSerializable(typeof(string[]))]
// ...existing attributes...
public partial class DocxodusJsonContext : JsonSerializerContext { }
```

- [ ] **Step 9.3: Build WASM and verify it compiles**

```bash
./scripts/build-wasm.sh 2>&1 | tail -15
```

Expected: clean build.

- [ ] **Step 9.4: Commit**

```bash
git add wasm/DocxodusWasm/
git commit -m "feat(wasm): expose textPreview on AnchorTargetDto + GetAnchorInfo/s bridge methods (#162)"
```

---

## Task 10: npm wrapper — add `textPreview` to types, expose `getAnchorInfo` / `getAnchorInfos`

**Files:**
- Modify: `npm/src/types.ts` — `MarkdownAnchorTarget` (line 85) and `AnchorTargetRef` (line 940). Add `AnchorInfo` type.
- Modify: `npm/src/session.ts` — add `getAnchorInfo` and `getAnchorInfos` methods.

- [ ] **Step 10.1: Update TypeScript types**

In `npm/src/types.ts` find `MarkdownAnchorTarget` around line 85 and add `textPreview`:

```typescript
export interface MarkdownAnchorTarget {
  id: string;
  kind: string;
  scope: string;
  unid: string;
  partUri: string;
  textPreview: string;   // NEW — first ~80 chars of element's flat text
}
```

Find `AnchorTargetRef` around line 940 and add the same field. (Read the surrounding interface to confirm the property naming pattern matches.)

Add a new exported interface near the other anchor types:

```typescript
/**
 * The shape returned by {@link DocxSession.getAnchorInfo}.
 * Use {@link MarkdownAnchorTarget} when iterating a full projection — it
 * includes the same fields plus `unid` and `partUri`.
 */
export interface AnchorInfo {
  id: string;
  kind: string;
  scope: string;
  textPreview: string;
}
```

- [ ] **Step 10.2: Add `DocxodusWasmExports` entries for the new bridge methods**

In `npm/src/types.ts`, locate the `DocxodusWasmExports` interface (search for the type that mirrors the bridge — it'll have method signatures like `ReplaceText`, `FindByLabel`, etc.). Add:

```typescript
export interface DocxodusWasmExports {
  // ... existing entries ...
  GetAnchorInfo(handle: number, anchorId: string): string;
  GetAnchorInfos(handle: number, anchorIdsJson: string): string;
}
```

- [ ] **Step 10.3: Add wrapper methods on `DocxSession`**

In `npm/src/session.ts`, add (near the other discovery helpers like `findByBookmark`):

```typescript
/**
 * Look up a single anchor's preview info — `{ id, kind, scope, textPreview }`.
 * Returns null when the anchor id is unknown.
 *
 * For iterating many anchors at once, prefer reading `textPreview` directly
 * off the {@link MarkdownProjection.anchorIndex} entries (cheaper — no extra
 * WASM round trip), or use {@link getAnchorInfos} for batched lookups.
 */
getAnchorInfo(anchorId: string): AnchorInfo | null {
  const raw = this.wasm.GetAnchorInfo(this.handle, anchorId);
  return JSON.parse(raw) as AnchorInfo | null;
}

/**
 * Bulk variant of {@link getAnchorInfo}: takes an array of anchor ids,
 * returns a record where each unknown id maps to `null`.
 */
getAnchorInfos(anchorIds: readonly string[]): Record<string, AnchorInfo | null> {
  const raw = this.wasm.GetAnchorInfos(this.handle, JSON.stringify(anchorIds));
  return JSON.parse(raw) as Record<string, AnchorInfo | null>;
}
```

Add `AnchorInfo` to the import block at the top of `session.ts` (search for the import statement that pulls `AnchorRef, AnchorTargetRef`).

- [ ] **Step 10.4: Build the npm package**

```bash
cd /home/jman/Code/Docxodus/npm && npm run build 2>&1 | tail -20
```

Expected: clean build (TypeScript compiles, esbuild bundles).

- [ ] **Step 10.5: Type-check**

```bash
cd /home/jman/Code/Docxodus/npm && npx tsc --noEmit 2>&1 | tail -10
```

Expected: no errors.

- [ ] **Step 10.6: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/src/
git commit -m "feat(npm): textPreview on MarkdownAnchorTarget, expose getAnchorInfo(s) (#162)"
```

---

## Task 11: Add Playwright integration test for textPreview surfacing in npm

**Files:**
- Create: `npm/tests/anchor-introspection.spec.ts`

- [ ] **Step 11.1: Write the spec**

Inspect existing specs to see the harness shape:

```bash
ls /home/jman/Code/Docxodus/npm/tests/ | head -20
head -40 /home/jman/Code/Docxodus/npm/tests/$(ls /home/jman/Code/Docxodus/npm/tests/ | grep -E "session|find" | head -1)
```

Then create `npm/tests/anchor-introspection.spec.ts` following the same harness pattern. The test should:

1. Load a fixture DOCX in the browser harness.
2. Project it and assert `projection.anchorIndex` entries have non-empty `textPreview` for body block kinds.
3. Call `session.getAnchorInfo(id)` and assert it returns the same `textPreview`.
4. Call `session.getAnchorInfos([id1, id2, 'unknown'])` and assert the map shape.
5. Assert no fn-kind anchor exists for boilerplate separators (if the fixture has footnotes).

Use the most similar existing spec as a template (likely `findByLabel.spec.ts` or `grep.spec.ts`).

- [ ] **Step 11.2: Run the new spec**

```bash
cd /home/jman/Code/Docxodus/npm && npx playwright test --grep "anchor-introspection" 2>&1 | tail -15
```

Expected: green. If the Playwright browser isn't installed: `npx playwright install chromium` first.

- [ ] **Step 11.3: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/tests/anchor-introspection.spec.ts
git commit -m "test(npm): anchor introspection — textPreview + bulk lookup (#162)"
```

---

## Task 12: Update documentation

**Files:**
- Modify: `CHANGELOG.md` — add entry under `[Unreleased]`.
- Modify: `docs/architecture/markdown_projection.md` — document the new `TextPreview` field.
- Modify: `docs/architecture/docx_mutation_api.md` — note the `GetAnchorInfos` addition.
- Modify: `~/Code/Docxodus-Agents/PITFALLS.md` — §7 (Word-reserved footnotes) gets a "fixed in #162" note.
- Modify: `~/Code/Docxodus-Agents/TOOL_SURFACE.md` — mention `textPreview` on `AnchorTarget`.

- [ ] **Step 12.1: CHANGELOG**

Open `CHANGELOG.md`, find the `## [Unreleased]` section, and add under `### Added`:

```markdown
### Added
- **Anchor introspection ergonomics — `TextPreview` on `AnchorTarget`, boilerplate footnote filter, `GetAnchorInfos` bulk lookup** (issue #162). `WmlToMarkdownConverter` now computes the first ~80 chars of each block element's flat text during projection and exposes it as `AnchorTarget.TextPreview` — agents no longer need an N-anchor walk via `session.GetAnchorInfo` to surface previews when iterating the `AnchorIndex`. Word-reserved `w:footnote`/`w:endnote` separators (`type="separator"` / `type="continuationSeparator"`) no longer appear in the projection's `AnchorIndex` (they were internal Word plumbing surfaced as un-deletable `fn:fn:*` anchors). New `DocxSession.GetAnchorInfos(IEnumerable<string>)` returns a dictionary mapping each requested id to its `AnchorInfo?` in a single pass; unknown ids map to `null`. WASM bridge surfaces `textPreview` on `MarkdownAnchorTargetDto` and adds `GetAnchorInfo` / `GetAnchorInfos` JSExports. npm wrapper: new `textPreview` field on `MarkdownAnchorTarget` and `AnchorTargetRef`; `session.getAnchorInfo(id)` and `session.getAnchorInfos(ids[])` methods. Tests: `MD040`–`MD041`, `DS220`–`DS221`, Playwright `anchor-introspection.spec.ts`.
```

- [ ] **Step 12.2: `docs/architecture/markdown_projection.md`**

Open the file and find the "Round-Trip: Anchor → XElement" section (around line 149). After the existing block describing `target.PartUri`/`target.ElementXPath`/`target.Unid`, add:

```markdown
The `AnchorTarget` also exposes a `TextPreview` field — the first ~80 characters
of the element's flat text — computed during projection. Agents iterating
`AnchorIndex` for a UI list or LLM context window can read previews directly
without re-walking each element via `session.GetAnchorInfo`.

Word-reserved footnote/endnote separators (`type="separator"` /
`type="continuationSeparator"`) are excluded from `AnchorIndex` — they're
structural plumbing for Word's separator-line rendering, have no editorial
content, and cannot be deleted. They do not appear in the projection text either.
```

- [ ] **Step 12.3: `docs/architecture/docx_mutation_api.md`**

Find the existing reference to `GetAnchorInfo` (search the file). Add a short note about `GetAnchorInfos`:

```markdown
For batched lookups (an agent that just enumerated 50 anchors and wants
previews for all of them), use `session.GetAnchorInfos(ids)` — a single pass
over the AnchorIndex instead of one walk per id. Returns
`IReadOnlyDictionary<string, AnchorInfo?>` — unknown ids map to null.
```

- [ ] **Step 12.4: Update the agent guide**

Open `~/Code/Docxodus-Agents/PITFALLS.md`, find §7 (Word-reserved footnotes). Add a note at the top of the section:

```markdown
## §7. Word-reserved footnotes can't be deleted

> **Fixed in Docxodus 6.1.0 (#162).** Word-reserved footnote separators no
> longer appear in `AnchorIndex` at all — you won't see them via
> `session.Project().AnchorIndex` anymore. This section is preserved for
> agents working against Docxodus 6.0.x or earlier.

[existing content remains]
```

Open `~/Code/Docxodus-Agents/TOOL_SURFACE.md`, find the "Anchor shapes you'll see" section, and add a brief mention near it (or earlier in the doc) that `AnchorTarget` now carries `TextPreview` directly:

```markdown
Each `AnchorTarget` in `projection.AnchorIndex` carries the first ~80 chars of
the element's flat text as `TextPreview` — read this directly when you just
need to identify or display an anchor; reach for `session.GetAnchorInfo(id)`
only if the agent doesn't already hold a projection.
```

- [ ] **Step 12.5: Commit**

```bash
cd /home/jman/Code/Docxodus
git add CHANGELOG.md docs/architecture/markdown_projection.md docs/architecture/docx_mutation_api.md
git commit -m "docs: anchor introspection — TextPreview field + GetAnchorInfos (#162)"

cd /home/jman/Code/Docxodus-Agents
git status                                  # confirm we're in a repo or just a dir
# If this isn't a git repo, skip the commit step — these are local-only agent docs.
```

Note: `~/Code/Docxodus-Agents` may not be a git repo. If not, just save the files; they're a local reference. The Docxodus commit is the canonical one.

---

## Task 13: Final verification — full test suite + WASM/npm round-trip

- [ ] **Step 13.1: Full .NET test suite**

```bash
cd /home/jman/Code/Docxodus
dotnet clean Docxodus.sln 2>&1 | tail -3   # required because build-wasm.sh left WASM-mode dlls
dotnet build Docxodus.sln 2>&1 | tail -5
dotnet test Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -10
```

Expected: all tests pass. If a pre-existing test broke, investigate — but don't change implementation without understanding the regression.

- [ ] **Step 13.2: npm build + Playwright**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npm test 2>&1 | tail -10
```

Expected: all Playwright tests pass, including the new `anchor-introspection.spec.ts`.

- [ ] **Step 13.3: Release build (warnings-as-errors check)**

```bash
cd /home/jman/Code/Docxodus
dotnet build -c Release Docxodus.sln 2>&1 | tail -10
```

Expected: clean release build. Per `Directory.Build.props`, Release fails on warnings; investigate any new warning.

- [ ] **Step 13.4: Push the branch and open PR**

```bash
git push -u origin feat/162-anchor-introspection
gh pr create --title "feat: anchor introspection ergonomics — TextPreview on AnchorTarget, filter reserved fns, bulk lookup" --body "$(cat <<'EOF'
Closes #162.

## Summary
- `TextPreview` on `AnchorTarget`, computed once during projection — no more N-walk `GetAnchorInfo` round-trips for previews.
- Word-reserved footnote/endnote separators no longer appear in `AnchorIndex` (they were un-deletable plumbing surfaced as `fn:fn:*` anchors).
- `DocxSession.GetAnchorInfos(ids)` for batched lookups; unknown ids → null.

## Test plan
- [ ] `MD040` — `AnchorTarget.TextPreview` populated for body block kinds.
- [ ] `MD041` — boilerplate fn separators excluded from `AnchorIndex`.
- [ ] `DS220` — `GetAnchorInfos` batched lookup; unknown ids map to null.
- [ ] `DS221` — `GetAnchorInfo` regression: same `TextPreview` as `AnchorTarget`.
- [ ] Playwright `anchor-introspection.spec.ts` — npm surface end-to-end.
- [ ] Full `dotnet test` passes.
- [ ] `npm run build && npm test` passes.
- [ ] Release build clean (warnings-as-errors).
EOF
)"
```

Expected: PR URL printed.

---

## Self-Review

**Spec coverage:**
- Issue #162 §1 (TextPreview on AnchorTarget): Tasks 2–3. ✓
- Issue #162 §2 (filter Word-reserved footnotes): Tasks 4–5. Option (a) chosen (skip during walk, not flag with `IsReserved`). ✓
- Issue #162 §3 (`GetAnchorInfos`): Tasks 6–7. ✓
- Acceptance criteria — "`AnchorTarget.TextPreview` populated for every block-level kind; empty acceptable for `sec`/`unk`": covered by Step 3.1 default + Step 3.3 populating from `ComputeTextPreview`. ✓
- Acceptance criteria — "Reserved footnote separators no longer appear in `AnchorIndex`": Task 5. ✓
- Acceptance criteria — "`GetAnchorInfos` returns results in O(parts-walked)": yes — single pass over IDs, each is an O(1) dictionary hit on `Project().AnchorIndex` via `FindAnchor`. ✓
- Acceptance criteria — "Docs updated": Task 12 covers both Docxodus repo docs and the agent guide. ✓
- Acceptance criteria — "Test: NVCA template, no `AnchorWrongKind` on `DeleteBlock` of any fn anchor": covered indirectly by Task 5's MD041 (boilerplate fns excluded from the index, so a loop over all fn anchors can't hit one). The literal NVCA-template assertion would be heavyweight; the synthetic fixture covers the same invariant.

**Placeholder scan:** No `TODO`, no `Similar to Task N`, no bare "Add error handling." Every step shows exact code or exact command.

**Type consistency:** `TextPreview` named identically across `AnchorTarget` (C#), `MarkdownAnchorTargetDto` (WASM C#), `MarkdownAnchorTarget`/`AnchorTargetRef`/`AnchorInfo` (TypeScript). `textPreview` (camelCase) in the JSON serialization and TypeScript surface — matches existing convention (e.g. `partUri`).

---

**Plan complete and saved to `docs/superpowers/plans/2026-05-25-issue-162-anchor-introspection.md`.**

Two execution options:

1. **Subagent-Driven (recommended)** — dispatch a fresh subagent per task, review between tasks. Better isolation; each task's subagent doesn't carry context from earlier tasks.
2. **Inline Execution** — execute tasks in this session using `executing-plans`. Batch execution with checkpoints; faster but tasks share context (mostly fine for this size).
