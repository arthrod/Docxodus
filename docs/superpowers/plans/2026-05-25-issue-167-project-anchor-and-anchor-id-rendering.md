# Issue #167 — `ProjectAnchor` + `AnchorIdRendering` Modes

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Land two LLM-friendliness improvements on the projection surface: `DocxSession.ProjectAnchor(anchorId, depth?)` for per-section/subtree projection that fits in smaller context windows, and `WmlToMarkdownConverterSettings.AnchorIdRendering` (`FullUnid` / `Abbreviated` / `Sequential`) for shorter anchor tokens that save 5-10% of projection-token budget.

**Architecture:** `ProjectAnchor` runs `Project()` (memoized) once, computes the in-range Unid set by walking siblings/descendants from the requested anchor, then post-filters the full markdown into a sub-projection. `AnchorIdRendering` adds an `AnchorIdMap` to `EmitContext` that all `{#…}` emission sites consult; the projection's `AnchorIndex` is dual-keyed (full + rendered) so lookups work via either form. `Anchor.Id` itself remains the canonical full identifier — only the rendered tokens and alias keys change.

**Tech Stack:** .NET 8, C#, Open XML SDK 3.x, xUnit. Shared `Docxodus.Internal/` core. WASM `[JSExport]` bridge + Python stdio NDJSON dispatcher. npm TypeScript wrapper. Playwright for browser tests.

**Branch:** `feat/167-project-anchor-and-rendering`

---

## Task 1: Create feature branch + commit plan

- [ ] **Step 1.1: Confirm clean tree on main**

```bash
cd /home/jman/Code/Docxodus
git status
git log --oneline -3
```

Expected: clean on `main`; most recent commit is `aaaaf75` (PR #179 merge) or later.

- [ ] **Step 1.2: Branch + commit the plan**

```bash
git checkout -b feat/167-project-anchor-and-rendering
git add docs/superpowers/plans/2026-05-25-issue-167-project-anchor-and-anchor-id-rendering.md
git commit -m "docs: plan for issue #167 (ProjectAnchor + AnchorIdRendering)"
```

---

## Task 2: `AnchorIdRendering` enum + `Abbreviated` mode (failing tests)

**Files:**
- Test: `Docxodus.Tests/WmlToMarkdownConverterTests.cs` — append `MD050`–`MD052`.

- [ ] **Step 2.1: Add three failing tests**

```csharp
[Fact]
public void MD050_AnchorIdRendering_FullUnid_IsExistingBehavior()
{
    // Default rendering — anchor tokens carry the full 32-char hex Unid.
    var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
    var wml = new WmlDocument("test.docx", bytes);
    var settings = new WmlToMarkdownConverterSettings();   // AnchorIdRendering defaults to FullUnid
    var projection = WmlToMarkdownConverter.Convert(wml, settings);

    // Every {#…} token uses the full 32-char hex form (matched by regex {#kind:scope:UNID}).
    var match = System.Text.RegularExpressions.Regex.Match(
        projection.Markdown, @"\{#[^:]+:[^:]+:([a-f0-9]+)\}");
    Assert.True(match.Success, "expected at least one anchor token in the projection");
    Assert.Equal(32, match.Groups[1].Length);
}

[Fact]
public void MD051_AnchorIdRendering_Abbreviated_UsesShortestUniquePrefixPerScope()
{
    // Abbreviated mode picks the shortest prefix per (kind, scope) bucket
    // that uniquely identifies each anchor, with a 4-char floor.
    var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
    var wml = new WmlDocument("test.docx", bytes);
    var settings = new WmlToMarkdownConverterSettings
    {
        AnchorIdRendering = AnchorIdRendering.Abbreviated,
    };
    var projection = WmlToMarkdownConverter.Convert(wml, settings);

    // All emitted tokens are < 32 chars in their unid portion, and >= 4.
    foreach (System.Text.RegularExpressions.Match m in
             System.Text.RegularExpressions.Regex.Matches(projection.Markdown, @"\{#[^:]+:[^:]+:([a-f0-9]+)\}"))
    {
        var unid = m.Groups[1].Value;
        Assert.True(unid.Length >= 4, $"abbreviation '{unid}' shorter than 4-char floor");
        Assert.True(unid.Length < 32, $"abbreviation '{unid}' wasn't actually abbreviated");
    }
}

[Fact]
public void MD052_AnchorIdRendering_Abbreviated_AnchorIndexHasDualKeys()
{
    // The AnchorIndex must contain entries keyed by BOTH the full Unid and
    // the abbreviated id, both pointing at the same AnchorTarget. Callers
    // can use whichever form they have in hand.
    var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
    var wml = new WmlDocument("test.docx", bytes);
    var settings = new WmlToMarkdownConverterSettings
    {
        AnchorIdRendering = AnchorIdRendering.Abbreviated,
    };
    var projection = WmlToMarkdownConverter.Convert(wml, settings);

    // Find any anchor — extract its full Unid from the underlying AnchorTarget,
    // and the abbreviated form from a token in the markdown.
    var firstTarget = projection.AnchorIndex.Values
        .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h");
    var fullKey = firstTarget.Anchor.Id;
    Assert.True(projection.AnchorIndex.ContainsKey(fullKey),
        $"full key '{fullKey}' missing from index");

    // The matching abbreviated key is somewhere in the markdown — extract any one.
    var abbreviatedToken = System.Text.RegularExpressions.Regex.Match(
        projection.Markdown, @"\{#([^:]+:[^:]+:[a-f0-9]+)\}");
    Assert.True(abbreviatedToken.Success);
    var abbreviatedKey = abbreviatedToken.Groups[1].Value;
    Assert.NotEqual(fullKey, abbreviatedKey);  // it's actually abbreviated, not the full Unid
    Assert.True(projection.AnchorIndex.ContainsKey(abbreviatedKey),
        $"abbreviated key '{abbreviatedKey}' missing from index");

    // Both keys resolve to AnchorTarget whose underlying Unid matches.
    var fromFull = projection.AnchorIndex[fullKey];
    var fromAbbreviated = projection.AnchorIndex[abbreviatedKey];
    Assert.Equal(fromFull.Unid, fromAbbreviated.Unid);
}
```

- [ ] **Step 2.2: Run, verify all three fail**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD050|FullyQualifiedName~MD051|FullyQualifiedName~MD052" 2>&1 | tail -15
```

Expected: build error — `AnchorIdRendering` doesn't exist yet.

---

## Task 3: `AnchorIdRendering` enum + `Abbreviated` mode (implementation)

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs` — add `AnchorIdRendering` enum, `Settings.AnchorIdRendering` property, `AnchorIdMap` private class, `EmitContext.AnchorIdMap` field, abbreviation computation in `BuildAnchorIndex`, route all `{#…}` emission sites through `AnchorIdMap`, dual-key the `AnchorIndex`.

- [ ] **Step 3.1: Add `AnchorIdRendering` enum**

In `Docxodus/WmlToMarkdownConverter.cs`, near `EmptyParagraphMode` (around line 125), add:

```csharp
/// <summary>How anchor ids are rendered inside <c>{#…}</c> tokens and keyed in
/// <see cref="MarkdownProjection.AnchorIndex"/>.</summary>
public enum AnchorIdRendering
{
    /// <summary>Full 32-char hex Unid (e.g. <c>{#h:body:a1b2c3d4e5f6789012345678901234ab}</c>).
    /// Default; matches the projection's existing behavior.</summary>
    FullUnid = 0,

    /// <summary>Shortest unique prefix per (kind, scope) bucket, with a 4-char
    /// floor (e.g. <c>{#h:body:a1b2}</c>). Saves 5-10% of projection-token
    /// budget for LLM consumption. The <see cref="MarkdownProjection.AnchorIndex"/>
    /// is dual-keyed — lookups by either full Unid or abbreviated form work.</summary>
    Abbreviated = 1,

    /// <summary>Sequential numeric ids per (kind, scope) bucket, in document
    /// order (e.g. <c>{#h:body:1} {#h:body:2}</c>). Maximally token-efficient
    /// for one-shot LLM contexts where anchor stability across sessions is
    /// unnecessary. These ids are NOT stable across <c>Project()</c> calls and
    /// must NOT be persisted; the dual-keyed <c>AnchorIndex</c> contains the
    /// numeric form too, but the canonical <see cref="Anchor.Id"/> on each
    /// <c>AnchorTarget</c> still carries the full Unid.</summary>
    Sequential = 2,
}
```

- [ ] **Step 3.2: Add `AnchorIdRendering` property on settings**

In `WmlToMarkdownConverterSettings` (around line 74), add:

```csharp
/// <summary>How anchor ids are rendered in <c>{#…}</c> tokens and keyed in
/// the resulting <see cref="MarkdownProjection.AnchorIndex"/>. Default
/// <see cref="AnchorIdRendering.FullUnid"/> matches legacy behavior.</summary>
public AnchorIdRendering AnchorIdRendering { get; set; } = AnchorIdRendering.FullUnid;
```

- [ ] **Step 3.3: Add private `AnchorIdMap` class**

Near `BuildAnchorIndex` (around line 280), add a private nested helper class:

```csharp
/// <summary>
/// Per-projection map from full Unid → rendered id. Used by every <c>{#…}</c>
/// emission site so the projection's anchor tokens consistently use the
/// caller-requested rendering. The <see cref="Render"/> method returns the full
/// Unid unchanged when the rendering is <see cref="AnchorIdRendering.FullUnid"/>
/// or when an unknown Unid is passed (defensive fallback).
/// </summary>
internal sealed class AnchorIdMap
{
    private readonly Dictionary<string, string> _map = new(StringComparer.Ordinal);
    public string Render(string fullUnid) => _map.TryGetValue(fullUnid, out var r) ? r : fullUnid;
    internal void Set(string fullUnid, string renderedUnid) => _map[fullUnid] = renderedUnid;
    internal IReadOnlyDictionary<string, string> Map => _map;
}
```

- [ ] **Step 3.4: Compute abbreviation in `BuildAnchorIndex`**

In `BuildAnchorIndex`, AFTER the main descendant-walk loop that populates `index`, BEFORE the return statement, add:

```csharp
// Build the AnchorIdMap based on the requested rendering mode.
var renderMap = new AnchorIdMap();
if (settings.AnchorIdRendering == AnchorIdRendering.Abbreviated)
{
    // Group anchors by (kind, scope). Within each group, find the shortest
    // prefix length n >= 4 such that every member's Unid[..n] is unique.
    foreach (var bucket in index.Values.GroupBy(t => (t.Anchor.Kind, t.Anchor.Scope)))
    {
        var members = bucket.ToList();
        if (members.Count == 0) continue;
        int n = 4;
        while (true)
        {
            var prefixes = new HashSet<string>(StringComparer.Ordinal);
            bool unique = true;
            foreach (var t in members)
            {
                var prefix = t.Unid.Length >= n ? t.Unid.Substring(0, n) : t.Unid;
                if (!prefixes.Add(prefix)) { unique = false; break; }
            }
            if (unique) break;
            n++;
            if (n >= 32) break;   // fall back to full Unid
        }
        foreach (var t in members)
        {
            var prefix = t.Unid.Length >= n ? t.Unid.Substring(0, n) : t.Unid;
            renderMap.Set(t.Unid, prefix);
        }
    }
}

// Dual-key the index: add alias entries with the rendered id substituted.
if (settings.AnchorIdRendering != AnchorIdRendering.FullUnid)
{
    var aliases = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
    foreach (var (key, target) in index)
    {
        var rendered = renderMap.Render(target.Unid);
        if (rendered == target.Unid) continue;
        var aliasKey = $"{target.Anchor.Kind}:{target.Anchor.Scope}:{rendered}";
        aliases[aliasKey] = target;
    }
    foreach (var (key, target) in aliases)
        index[key] = target;
}
```

Then **return both** `(index, scopes, renderMap)` from `BuildAnchorIndex` — change the tuple return type. The downstream call site in `Convert` will pass `renderMap` into the `EmitContext`.

Updated signature:

```csharp
private static (IReadOnlyDictionary<string, AnchorTarget> Index, List<ScopeInfo> Scopes, AnchorIdMap RenderMap)
    BuildAnchorIndex(WordprocessingDocument doc, WmlToMarkdownConverterSettings settings)
```

- [ ] **Step 3.5: Thread `AnchorIdMap` into `EmitContext`**

In `EmitContext` (around line 415), add a required field:

```csharp
required public AnchorIdMap AnchorIdMap { get; init; }
```

In `Convert(WordprocessingDocument, WmlToMarkdownConverterSettings)` (around line 261), update the destructure and the `EmitContext` construction inside `EmitMarkdown`'s call site:

```csharp
var (index, scopes, renderMap) = BuildAnchorIndex(document, settings);
var markdown = EmitMarkdown(document, settings, scopes, renderMap);
```

And `EmitMarkdown` signature gets a new parameter:

```csharp
private static string EmitMarkdown(WordprocessingDocument document, WmlToMarkdownConverterSettings settings,
    List<ScopeInfo> scopes, AnchorIdMap renderMap)
```

Inside, where it constructs `EmitContext`, pass `AnchorIdMap = renderMap`.

- [ ] **Step 3.6: Route `{#…}` emission sites through `AnchorIdMap`**

There are several emission sites:

1. **`AnchorPrefix`** (line 736) — central helper:

```csharp
private static string AnchorPrefix(XElement el, EmitContext ctx)
{
    if (ctx.Settings.AnchorMode == AnchorRenderMode.None) return string.Empty;
    var kind = KindFor(el) ?? "unk";
    var unid = (string?)el.Attribute(PtOpenXml.Unid) ?? "0";
    var rendered = ctx.AnchorIdMap.Render(unid);   // NEW
    return $"{{#{kind}:{ctx.Scope}:{rendered}}} ";
}
```

2. **`{#sec:...}` emission** (around line 729): apply the same `ctx.AnchorIdMap.Render(unid)` substitution before formatting the string.

3. **`{#cmt:cmt:{unid}}` emission in comments section** (around line 587): same substitution.

4. **`[^fn-...]` footnote reference emission**: check whether this uses `ShortUnid(unid)` or another helper. If it uses the full Unid, route through `ctx.AnchorIdMap.Render(unid)` similarly.

5. **`{#img:...}` / `{#unk:...}` and any other inline anchor markers**: grep for `{#` in the file and verify each uses the map.

Grep to find all sites:

```bash
grep -nE "\"\\{#|\\$\"\\{#" /home/jman/Code/Docxodus/Docxodus/WmlToMarkdownConverter.cs | head -20
```

Update each.

- [ ] **Step 3.7: Run the three new tests**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD050|FullyQualifiedName~MD051|FullyQualifiedName~MD052" 2>&1 | tail -10
```

Expected: 3/3 pass.

- [ ] **Step 3.8: Run the full WmlToMarkdownConverter + DocxSession regression**

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~WmlToMarkdownConverterTests|FullyQualifiedName~DocxSessionTests" 2>&1 | tail -5
```

Expected: all green. If any test asserts on a specific anchor token containing 32-char hex, the default `FullUnid` mode preserves that.

- [ ] **Step 3.9: Commit**

```bash
git add Docxodus/WmlToMarkdownConverter.cs Docxodus.Tests/WmlToMarkdownConverterTests.cs
git commit -m "feat(markdown-projection): AnchorIdRendering enum + Abbreviated mode (#167)"
```

---

## Task 4: `Sequential` rendering mode (failing tests)

**Files:**
- Test: `Docxodus.Tests/WmlToMarkdownConverterTests.cs` — append `MD053`–`MD054`.

- [ ] **Step 4.1: Add the two tests**

```csharp
[Fact]
public void MD053_AnchorIdRendering_Sequential_NumbersPerScopeKindBucket()
{
    var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
    var wml = new WmlDocument("test.docx", bytes);
    var settings = new WmlToMarkdownConverterSettings
    {
        AnchorIdRendering = AnchorIdRendering.Sequential,
    };
    var projection = WmlToMarkdownConverter.Convert(wml, settings);

    // Each (kind, scope) bucket starts at 1 and increments per anchor in document order.
    // BuildDS001 has 2 body paragraphs → tokens "{#p:body:1}" and "{#p:body:2}".
    Assert.Contains("{#p:body:1}", projection.Markdown);
    Assert.Contains("{#p:body:2}", projection.Markdown);
}

[Fact]
public void MD054_AnchorIdRendering_Sequential_AnchorIndexHasDualKeys()
{
    var bytes = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
    var wml = new WmlDocument("test.docx", bytes);
    var settings = new WmlToMarkdownConverterSettings
    {
        AnchorIdRendering = AnchorIdRendering.Sequential,
    };
    var projection = WmlToMarkdownConverter.Convert(wml, settings);

    // The full key still resolves; AND the sequential alias key resolves;
    // and both point at the same target.
    var firstTarget = projection.AnchorIndex.Values
        .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind == "p");
    Assert.True(projection.AnchorIndex.ContainsKey(firstTarget.Anchor.Id));
    Assert.True(projection.AnchorIndex.ContainsKey("p:body:1"));
    Assert.Same(firstTarget, projection.AnchorIndex["p:body:1"]);
}
```

- [ ] **Step 4.2: Run, verify both fail**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD053|FullyQualifiedName~MD054" 2>&1 | tail -10
```

Expected: tests run but the assertions fail because `Sequential` mode isn't implemented yet (the dictionary fallback returns the full Unid in `Render`).

---

## Task 5: `Sequential` rendering mode (implementation)

**Files:**
- Modify: `Docxodus/WmlToMarkdownConverter.cs::BuildAnchorIndex` — add Sequential branch.

- [ ] **Step 5.1: Extend `BuildAnchorIndex` with the Sequential branch**

In `BuildAnchorIndex`, in the rendering-map computation block from Step 3.4, add an else-if branch for `Sequential`:

```csharp
else if (settings.AnchorIdRendering == AnchorIdRendering.Sequential)
{
    // Assign 1-based numeric ids per (kind, scope) bucket in document order
    // (which is the natural order of index.Values since Dictionary preserves
    // insertion order and BuildAnchorIndex inserts in descendant-walk order).
    var counters = new Dictionary<(string Kind, string Scope), int>();
    foreach (var t in index.Values)
    {
        var bucket = (t.Anchor.Kind, t.Anchor.Scope);
        if (!counters.TryGetValue(bucket, out var n)) n = 0;
        n++;
        counters[bucket] = n;
        renderMap.Set(t.Unid, n.ToString(System.Globalization.CultureInfo.InvariantCulture));
    }
}
```

- [ ] **Step 5.2: Run the new tests + regression**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~MD053|FullyQualifiedName~MD054" 2>&1 | tail -10
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~WmlToMarkdownConverterTests|FullyQualifiedName~DocxSessionTests" 2>&1 | tail -5
```

Expected: 2/2 new tests pass + full regression green.

- [ ] **Step 5.3: Commit**

```bash
git add Docxodus/WmlToMarkdownConverter.cs Docxodus.Tests/WmlToMarkdownConverterTests.cs
git commit -m "feat(markdown-projection): AnchorIdRendering Sequential mode (#167)"
```

---

## Task 6: `ProjectAnchor` + `ProjectionDepth` enum (failing tests)

**Files:**
- Test: `Docxodus.Tests/DocxSessionTests.cs` — append `DS295`–`DS298`.

- [ ] **Step 6.1: Add four failing tests**

Tests reuse `BuildDocWithHeadingSections` from #165 (H1 → p → H2 → p → H1 → p, 6 blocks total).

```csharp
[Fact]
public void DS295_ProjectAnchor_OnParagraph_ReturnsJustThatParagraph()
{
    using var session = new DocxSession(BuildDocWithHeadingSections());
    var firstPara = session.Project().AnchorIndex.Values
        .First(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "para 1.1");

    var sub = session.ProjectAnchor(firstPara.Anchor.Id);
    Assert.Contains("para 1.1", sub.Markdown);
    Assert.DoesNotContain("Section One.A", sub.Markdown);
    Assert.DoesNotContain("Section Two", sub.Markdown);
}

[Fact]
public void DS296_ProjectAnchor_OnHeading_ReturnsTheWholeSection()
{
    using var session = new DocxSession(BuildDocWithHeadingSections());
    var sectionOne = session.Project().AnchorIndex.Values
        .First(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section One");

    var sub = session.ProjectAnchor(sectionOne.Anchor.Id);
    // The Heading1 section runs from "Section One" through everything before the next Heading1.
    Assert.Contains("Section One", sub.Markdown);
    Assert.Contains("para 1.1", sub.Markdown);
    Assert.Contains("Section One.A", sub.Markdown);
    Assert.Contains("para 1.A.1", sub.Markdown);
    // …but does NOT include "Section Two" (the next Heading1 boundary).
    Assert.DoesNotContain("Section Two", sub.Markdown);
    Assert.DoesNotContain("para 2.1", sub.Markdown);
}

[Fact]
public void DS297_ProjectAnchor_OnLastHeading_ExtendsToEndOfParent()
{
    using var session = new DocxSession(BuildDocWithHeadingSections());
    var sectionTwo = session.Project().AnchorIndex.Values
        .First(t => session.GetAnchorInfo(t.Anchor.Id)?.TextPreview == "Section Two");

    var sub = session.ProjectAnchor(sectionTwo.Anchor.Id);
    Assert.Contains("Section Two", sub.Markdown);
    Assert.Contains("para 2.1", sub.Markdown);
    Assert.DoesNotContain("Section One", sub.Markdown);
}

[Fact]
public void DS298_ProjectAnchor_UnknownAnchorThrows()
{
    using var session = new DocxSession(BuildDocWithHeadingSections());
    var ex = Assert.Throws<InvalidOperationException>(() =>
        session.ProjectAnchor("p:body:0000000000000000ffffffffffffffff"));
    Assert.Contains("not found", ex.Message);
}
```

- [ ] **Step 6.2: Run, verify all fail**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS295|FullyQualifiedName~DS296|FullyQualifiedName~DS297|FullyQualifiedName~DS298" 2>&1 | tail -15
```

Expected: build error — `'DocxSession' does not contain a definition for 'ProjectAnchor'`.

---

## Task 7: `ProjectAnchor` + `ProjectionDepth` enum (implementation)

**Files:**
- Modify: `Docxodus/DocxSession.cs` — add `ProjectionDepth` enum, `ProjectAnchor` method.

- [ ] **Step 7.1: Add `ProjectionDepth` enum**

Near the existing `DiffFormat` / `EditSummary` records (around line 340), add:

```csharp
/// <summary>How far below the target anchor to include in <see cref="DocxSession.ProjectAnchor"/>.</summary>
public enum ProjectionDepth
{
    /// <summary>Just the target block itself (its anchor + its own text). For headings,
    /// returns only the heading paragraph, not the section under it.</summary>
    SelfOnly = 0,

    /// <summary>Self + descendants. Most useful for <c>tbl</c> anchors (returns the whole
    /// table); for paragraphs it's the same as <see cref="SelfOnly"/>.</summary>
    Subtree = 1,

    /// <summary>Self + descendants + following siblings up to (but not including) the
    /// next sibling at the same or higher heading level. For non-heading anchors,
    /// equivalent to <see cref="Subtree"/>. This is the dominant "give me this section"
    /// case for headings and is the default.</summary>
    SubtreeAndFollowingSiblings = 2,
}
```

- [ ] **Step 7.2: Add `ProjectAnchor` method**

Near the existing `Project()` method (around line 459), add:

```csharp
/// <summary>
/// Project a sub-region of the document anchored at <paramref name="anchorId"/>.
/// Returns a <see cref="MarkdownProjection"/> whose <c>Markdown</c> contains only
/// the blocks in scope (per <paramref name="depth"/>) and whose <c>AnchorIndex</c>
/// is filtered to those blocks plus their descendants.
/// </summary>
/// <param name="anchorId">The anchor to project. Must exist in the current
/// <see cref="Project"/>'s AnchorIndex.</param>
/// <param name="depth">How far below the target to include. Default
/// <see cref="ProjectionDepth.SubtreeAndFollowingSiblings"/> — for headings, returns
/// the full section bounded by the next same-or-higher heading.</param>
/// <returns>A <see cref="MarkdownProjection"/> scoped to the requested region.</returns>
/// <exception cref="InvalidOperationException">If <paramref name="anchorId"/> isn't in the AnchorIndex.</exception>
public MarkdownProjection ProjectAnchor(
    string anchorId,
    ProjectionDepth depth = ProjectionDepth.SubtreeAndFollowingSiblings)
{
    ThrowIfDisposed();
    ArgumentNullException.ThrowIfNull(anchorId);

    var fullProjection = Project();
    var target = FindAnchor(anchorId)
        ?? throw new InvalidOperationException($"anchor not found: {anchorId}");

    var startElement = target.Resolve(_doc!)
        ?? throw new InvalidOperationException($"anchor element resolved null: {anchorId}");

    // Compute the set of Unids in scope.
    var inRange = new HashSet<string>(StringComparer.Ordinal);
    CollectUnids(startElement, inRange);

    if (depth == ProjectionDepth.SubtreeAndFollowingSiblings && target.Anchor.Kind == "h")
    {
        // For headings, also include forward siblings up to next same-or-higher heading.
        int targetLevel = WmlToMarkdownConverter.HeadingLevel(startElement);
        foreach (var sibling in startElement.ElementsAfterSelf())
        {
            if (sibling.Name == W.p
                && WmlToMarkdownConverter.IsHeading(sibling)
                && WmlToMarkdownConverter.HeadingLevel(sibling) <= targetLevel)
            {
                break;  // hit the section boundary
            }
            CollectUnids(sibling, inRange);
        }
    }
    else if (depth == ProjectionDepth.Subtree)
    {
        // CollectUnids already added self + descendants; nothing more to do.
    }
    // SelfOnly: CollectUnids was wrong — we should NOT have added descendants for that mode.
    // Re-derive for SelfOnly here:
    if (depth == ProjectionDepth.SelfOnly)
    {
        inRange.Clear();
        var selfUnid = (string?)startElement.Attribute(PtOpenXml.Unid);
        if (selfUnid is not null) inRange.Add(selfUnid);
    }

    // Filter the full markdown to blocks whose anchor token is in-range.
    // Blocks are separated by blank lines; each in-range block starts with {#kind:scope:unid}.
    var sb = new StringBuilder();
    foreach (var block in fullProjection.Markdown.Split("\n\n"))
    {
        var match = System.Text.RegularExpressions.Regex.Match(block, @"\{#[^:]+:[^:]+:([^\s}]+)\}");
        if (!match.Success) continue;  // skip scope markers / dividers / etc.
        var rendered = match.Groups[1].Value;
        // The rendered id might be the abbreviated or sequential form — translate back to full Unid.
        // For FullUnid mode this is the same string. For Abbreviated/Sequential the AnchorIndex
        // has alias keys; we don't need translation because we'll check both the full and rendered
        // forms via the index.
        if (TryResolveToUnid(rendered, match, fullProjection, out var fullUnid)
            && inRange.Contains(fullUnid))
        {
            sb.Append(block).Append("\n\n");
        }
    }

    // Filter the AnchorIndex too — keep only entries whose Unid is in scope.
    var filteredIndex = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
    foreach (var (key, value) in fullProjection.AnchorIndex)
    {
        if (inRange.Contains(value.Unid))
            filteredIndex[key] = value;
    }

    return new MarkdownProjection
    {
        Markdown = sb.ToString().TrimEnd('\n'),
        AnchorIndex = filteredIndex,
    };
}

private static void CollectUnids(XElement el, HashSet<string> sink)
{
    var unid = (string?)el.Attribute(PtOpenXml.Unid);
    if (unid is not null) sink.Add(unid);
    foreach (var d in el.Descendants())
    {
        var dUnid = (string?)d.Attribute(PtOpenXml.Unid);
        if (dUnid is not null) sink.Add(dUnid);
    }
}

/// <summary>
/// Resolve a rendered anchor id (full Unid, abbreviation, or sequential) back to
/// the underlying full Unid by looking it up in the projection's AnchorIndex
/// (which is dual-keyed when AnchorIdRendering is Abbreviated/Sequential).
/// </summary>
private static bool TryResolveToUnid(
    string rendered,
    System.Text.RegularExpressions.Match match,
    MarkdownProjection projection,
    out string fullUnid)
{
    // Try the full key (kind:scope:rendered) — works for FullUnid and as alias for others.
    var fullKey = match.Value[2..^1];  // strip {# and }
    if (projection.AnchorIndex.TryGetValue(fullKey, out var target))
    {
        fullUnid = target.Unid;
        return true;
    }
    fullUnid = rendered;
    return false;
}
```

- [ ] **Step 7.3: Run new tests + regression**

```bash
cd /home/jman/Code/Docxodus
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DS295|FullyQualifiedName~DS296|FullyQualifiedName~DS297|FullyQualifiedName~DS298" 2>&1 | tail -10
dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter "FullyQualifiedName~DocxSessionTests" 2>&1 | tail -5
```

Expected: 4/4 new tests pass + full regression green.

- [ ] **Step 7.4: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): ProjectAnchor for per-section / subtree projection (#167)"
```

---

## Task 8: Shared core + WASM bridge + Python stdio

**Files:**
- Modify: `Docxodus/Internal/DocxSessionOps.cs` — add `ProjectAnchor` op.
- Modify: `Docxodus/Internal/DocxSessionJson.cs::ParseProjectionSettings` — parse `anchorIdRendering` from settings JSON.
- Modify: `wasm/DocxodusWasm/DocxSessionBridge.cs` — add `ProjectAnchor` JSExport.
- Modify: `tools/python-host/Dispatcher.cs` — add `project_anchor` op.

- [ ] **Step 8.1: Add `ProjectAnchor` to `DocxSessionOps`**

In `Docxodus/Internal/DocxSessionOps.cs`, near the existing `Project` op (around line 28), add:

```csharp
public static string ProjectAnchor(int handle, string anchorId, ProjectionDepth depth) =>
    DocxSessionJson.SerializeProjection(SessionRegistry.Get(handle).ProjectAnchor(anchorId, depth));
```

- [ ] **Step 8.2: Add `anchorIdRendering` to `ParseProjectionSettings`**

In `Docxodus/Internal/DocxSessionJson.cs`, find the existing parsing of `WmlToMarkdownConverterSettings` from JSON (search for `ParseProjectionSettings` or wherever `ProjectionSettings` is materialized from the session settings JSON). Add a parse line:

```csharp
if (root.TryGetProperty("anchorIdRendering", out var air) && air.ValueKind == JsonValueKind.Number)
    settings.AnchorIdRendering = (AnchorIdRendering)air.GetInt32();
```

If `WmlToMarkdownConverterSettings` is built piecewise from the JSON in the bridge, mirror that pattern.

- [ ] **Step 8.3: Add `ProjectAnchor` JSExport**

In `wasm/DocxodusWasm/DocxSessionBridge.cs`, near the existing `Project` JSExport (search for `public static string Project`), add:

```csharp
/// <summary>
/// Bridge for <see cref="DocxSession.ProjectAnchor"/>. <paramref name="depth"/>
/// uses the numeric layout of <see cref="ProjectionDepth"/> (SelfOnly=0,
/// Subtree=1, SubtreeAndFollowingSiblings=2). Returns a JSON object with
/// the standard MarkdownProjection shape (markdown + anchorIndex).
/// </summary>
[JSExport]
public static string ProjectAnchor(int h, string anchorId, int depth) =>
    DocxSessionOps.ProjectAnchor(h, anchorId, (ProjectionDepth)depth);
```

- [ ] **Step 8.4: Add `project_anchor` op to Python dispatcher**

In `tools/python-host/Dispatcher.cs`, near `"project"` op, add:

```csharp
"project_anchor" => DocxSessionOps.ProjectAnchor(
    Handle(args), Str(args, "anchorId"),
    (ProjectionDepth)IntOptional(args, "depth", 2)),
```

- [ ] **Step 8.5: Build WASM target**

```bash
cd /home/jman/Code/Docxodus
./scripts/build-wasm.sh 2>&1 | tail -5
```

Expected: clean build.

- [ ] **Step 8.6: Commit**

```bash
git add Docxodus/Internal/DocxSessionOps.cs Docxodus/Internal/DocxSessionJson.cs wasm/DocxodusWasm/DocxSessionBridge.cs tools/python-host/Dispatcher.cs
git commit -m "feat(bridge): wire ProjectAnchor + anchorIdRendering through ops, WASM, Python stdio (#167)"
```

---

## Task 9: npm wrapper

**Files:**
- Modify: `npm/src/types.ts` — add `ProjectionDepth` / `AnchorIdRendering` const enums; extend `MarkdownProjectionSettings` (or `DocxSessionSettings.ProjectionSettings` — wherever projection settings live on the TS side) with `anchorIdRendering`; declare bridge method.
- Modify: `npm/src/session.ts` — add `projectAnchor` method.
- Modify: `npm/src/index.ts` — re-export the new enums.

- [ ] **Step 9.1: Add TS const enums**

In `npm/src/types.ts`, near `PlaceholderKinds`:

```typescript
/**
 * Numeric flag layout matching the .NET `ProjectionDepth` enum.
 */
export const ProjectionDepth = {
  SelfOnly: 0,
  Subtree: 1,
  SubtreeAndFollowingSiblings: 2,
} as const;

/**
 * Numeric flag layout matching the .NET `AnchorIdRendering` enum.
 *
 * - `FullUnid` (default) — 32-char hex ids in `{#…}` tokens.
 * - `Abbreviated` — shortest unique prefix per (kind, scope) bucket, 4-char floor.
 * - `Sequential` — `1, 2, 3, …` per bucket; NOT stable across `project()` calls.
 */
export const AnchorIdRendering = {
  FullUnid: 0,
  Abbreviated: 1,
  Sequential: 2,
} as const;
```

- [ ] **Step 9.2: Extend projection-settings TS type with `anchorIdRendering`**

Find the TS-side projection settings type (likely on `DocxSessionSettings.projectionSettings` or a sibling interface). Add:

```typescript
anchorIdRendering?: number;
```

- [ ] **Step 9.3: Declare `ProjectAnchor` bridge method**

In `npm/src/types.ts`, find `DocxodusWasmExports.DocxSessionBridge`. Add:

```typescript
ProjectAnchor: (handle: number, anchorId: string, depth: number) => string;
```

- [ ] **Step 9.4: Add `projectAnchor` wrapper method on `DocxSession`**

In `npm/src/session.ts`, near `project()`, add:

```typescript
/**
 * Project a sub-region of the document anchored at `anchorId`. Returns a
 * `MarkdownProjection`-shaped object filtered to the anchor's section/subtree.
 *
 * @param anchorId The anchor to project. Must exist in the current `project()`'s AnchorIndex.
 * @param depth How far below the target to include. Default `SubtreeAndFollowingSiblings`
 *              returns the full heading section. See {@link ProjectionDepth}.
 *
 * @see docs/architecture/markdown_projection.md#projectanchor
 */
projectAnchor(anchorId: string, depth: number = ProjectionDepth.SubtreeAndFollowingSiblings): DocxSessionProjection {
  return JSON.parse(this.wasm.ProjectAnchor(this.handle, anchorId, depth)) as DocxSessionProjection;
}
```

Add `ProjectionDepth` to the value imports at the top of `session.ts`.

- [ ] **Step 9.5: Re-export the new enums from `index.ts`**

In `npm/src/index.ts`, add:

```typescript
export { ProjectionDepth, AnchorIdRendering } from "./types.js";
```

- [ ] **Step 9.6: Build and type-check**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npx tsc --noEmit 2>&1 | tail -5
```

Expected: both clean.

- [ ] **Step 9.7: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/src/
git commit -m "feat(npm): expose projectAnchor + AnchorIdRendering / ProjectionDepth enums (#167)"
```

---

## Task 10: Playwright integration test

**Files:**
- Create: `npm/tests/project-anchor-and-rendering.spec.ts`

- [ ] **Step 10.1: Write the spec**

Mirror the harness pattern from `npm/tests/edit-summary-and-diff.spec.ts` or `delete-range-section.spec.ts`. Five assertion areas:

1. **`projectAnchor` on a paragraph returns just that paragraph.** Load a multi-block fixture; project the first body paragraph; assert the markdown contains the paragraph's text and not the others.
2. **`projectAnchor` on a heading returns its section.** Heading-sections fixture; project the first H1; assert the section's content is included but the next H1 isn't.
3. **`projectAnchor` on last heading extends to end.** Project the last H1; assert content after it survives, content before doesn't.
4. **`AnchorIdRendering.Abbreviated` produces shorter tokens.** Open a session with `projectionSettings.anchorIdRendering = 1`; call `project()`; assert all `{#…}` tokens have unid portions < 32 chars and ≥ 4 chars.
5. **`AnchorIdRendering.Sequential` produces `1, 2, 3` per bucket.** Open with `anchorIdRendering = 2`; assert the markdown contains `{#p:body:1}` and `{#p:body:2}`.

For assertions 1-3, reuse fixture bytes from `delete-range-section.spec.ts` (the headings fixture). For 4-5, the simpler 5-paragraph fixture from `fill-placeholders.spec.ts` works.

Use the raw `DocxSessionBridge` per the existing-spec convention.

- [ ] **Step 10.2: Run the spec**

```bash
cd /home/jman/Code/Docxodus/npm
npm test -- --grep "project-anchor-and-rendering" 2>&1 | tail -15
```

Expected: 5 passing.

- [ ] **Step 10.3: Commit**

```bash
cd /home/jman/Code/Docxodus
git add npm/tests/project-anchor-and-rendering.spec.ts
git commit -m "test(npm): ProjectAnchor + AnchorIdRendering integration (#167)"
```

---

## Task 11: Documentation

**Files:**
- `CHANGELOG.md`
- `docs/architecture/markdown_projection.md`
- `~/Code/Docxodus-Agents/TOOL_SURFACE.md` (out-of-repo)

- [ ] **Step 11.1: CHANGELOG entry**

Under `## [Unreleased]` / `### Added`:

```markdown
- **`ProjectAnchor` + `AnchorIdRendering` for LLM-friendly projection** (issue #167). `DocxSession.ProjectAnchor(anchorId, depth?)` returns a `MarkdownProjection` scoped to the target anchor's section/subtree — the dominant case is "project just this heading's section" (default `ProjectionDepth.SubtreeAndFollowingSiblings`, walks forward siblings until the next same-or-higher heading). For paragraphs returns just the paragraph; for tables returns the whole table; for non-heading blocks with `SubtreeAndFollowingSiblings` falls back to `Subtree`. Filters the `AnchorIndex` to in-scope anchors so callers can keep walking. New `WmlToMarkdownConverterSettings.AnchorIdRendering` enum offers three rendering modes for the `{#…}` tokens: `FullUnid` (default, 32-char hex — matches legacy behavior); `Abbreviated` (shortest unique prefix per (kind, scope) bucket with 4-char floor — saves 5-10% of projection-token budget for LLM contexts); `Sequential` (`1, 2, 3, …` per bucket in document order — maximally token-efficient, NOT stable across `project()` calls, opt-in only). In Abbreviated/Sequential modes the `AnchorIndex` is dual-keyed so callers can look up by either the full Unid id or the rendered alias — both keys resolve to the same `AnchorTarget`. `Anchor.Id` itself remains the canonical full Unid. Shared `DocxSessionOps.ProjectAnchor` propagates to WASM bridge + Python stdio (`project_anchor` op). npm wrapper: `session.projectAnchor(anchorId, depth?)` + `ProjectionDepth` / `AnchorIdRendering` const enums re-exported. Tests: `MD050`–`MD054`, `DS295`–`DS298`, Playwright `project-anchor-and-rendering.spec.ts`.
```

- [ ] **Step 11.2: `docs/architecture/markdown_projection.md`**

Find a sensible location (likely after the section discussing the `{#kind:scope:unid}` format or near the `WmlToMarkdownConverterSettings` reference). Add two new subsections:

```markdown
## Anchor id rendering modes

`WmlToMarkdownConverterSettings.AnchorIdRendering` controls the inside of every
`{#…}` token in the emitted markdown AND adds alias keys to the
`AnchorIndex` so lookups work with either form.

| Mode | Example token | Use when |
|---|---|---|
| `FullUnid` (default) | `{#h:body:a1b2c3d4e5f6789012345678901234ab}` | API consumers that need cross-session stability; everything in v1 of the projector |
| `Abbreviated` | `{#h:body:a1b2}` | LLM contexts where the 32-char hex is wasted token-budget; the 4-char floor is enforced; collision auto-expands |
| `Sequential` | `{#h:body:1}` | One-shot LLM contexts; ids are NOT stable across `Project()` calls and must not be persisted |

`Anchor.Id` on every `AnchorTarget` remains the canonical full Unid id — only
the rendered tokens in the markdown and the alias keys in `AnchorIndex` change.

## `ProjectAnchor` — per-anchor sub-projection

`session.ProjectAnchor(anchorId, depth?)` returns a `MarkdownProjection` scoped
to one anchor's section / subtree. Useful for LLM contexts where the full
document's projection is too large.

`ProjectionDepth` controls how far below the target to include:

- `SelfOnly` — just the target's own block. For headings, returns only the heading paragraph.
- `Subtree` — self + descendants. Right for `tbl` anchors; for paragraphs identical to `SelfOnly`.
- `SubtreeAndFollowingSiblings` (default) — self + descendants + forward siblings up to (but not including) the next same-or-higher heading. The natural "give me this section" mode.

For non-heading anchors with `SubtreeAndFollowingSiblings`, the behavior degrades to `Subtree`.
```

- [ ] **Step 11.3: Update agent guide — `TOOL_SURFACE.md`**

In `~/Code/Docxodus-Agents/TOOL_SURFACE.md`, find the "Lifecycle" table at the top. Add a row right after `session.Project()`:

```markdown
| Get a sub-projection for one anchor's section/subtree | `session.ProjectAnchor(anchorId, depth?)` — returns a `MarkdownProjection` scoped to the requested region |
```

Add a small "Token-budget tip" callout near the lifecycle section:

```markdown
**Tip — LLM token budget.** When feeding the projection to an LLM, set
`Settings.ProjectionSettings.AnchorIdRendering = AnchorIdRendering.Abbreviated`
to shrink anchor tokens by 5-10% (or `Sequential` for one-shot contexts where
the ids don't need to round-trip across calls).
```

- [ ] **Step 11.4: Commit**

```bash
cd /home/jman/Code/Docxodus
git add CHANGELOG.md docs/architecture/markdown_projection.md
git commit -m "docs: ProjectAnchor + AnchorIdRendering (#167)"
```

`~/Code/Docxodus-Agents/` is NOT a git repo — files saved in place.

---

## Task 12: Final verification + PR

- [ ] **Step 12.1: Full .NET test suite (clean rebuild)**

```bash
cd /home/jman/Code/Docxodus
rm -rf Docxodus/bin Docxodus/obj Docxodus.Tests/bin Docxodus.Tests/obj
dotnet build Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -5
dotnet test Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -10
```

Expected: 0 failed.

- [ ] **Step 12.2: npm build + Playwright**

```bash
cd /home/jman/Code/Docxodus/npm
npm run build 2>&1 | tail -5
npm test 2>&1 | tail -15
```

Expected: all green.

- [ ] **Step 12.3: Release build**

```bash
cd /home/jman/Code/Docxodus
rm -rf Docxodus/bin Docxodus/obj Docxodus.Tests/bin Docxodus.Tests/obj
dotnet build -c Release Docxodus/Docxodus.csproj 2>&1 | tail -5
dotnet build -c Release Docxodus.Tests/Docxodus.Tests.csproj 2>&1 | tail -5
```

Expected: clean (modulo pre-existing #173).

- [ ] **Step 12.4: Push and open PR**

```bash
git push -u origin feat/167-project-anchor-and-rendering 2>&1 | tail -5

gh pr create --title "feat: ProjectAnchor + AnchorIdRendering modes for LLM-friendly projection" --body "$(cat <<'EOF'
Closes #167.

## Summary
- `DocxSession.ProjectAnchor(anchorId, depth?)` — projects a sub-region of the document. Default `SubtreeAndFollowingSiblings` gives the full section under a heading; for paragraphs returns just the paragraph; for tables returns the whole table.
- `WmlToMarkdownConverterSettings.AnchorIdRendering` — three rendering modes for `{#…}` tokens:
  - `FullUnid` (default) — legacy 32-char hex.
  - `Abbreviated` — shortest unique prefix per (kind, scope) bucket, 4-char floor. Saves 5-10% token budget.
  - `Sequential` — `1, 2, 3, …` per bucket. Maximally token-efficient; NOT stable across `Project()` calls.
- `AnchorIndex` is dual-keyed in non-FullUnid modes — lookups by either full Unid id or rendered alias resolve to the same `AnchorTarget`.
- `Anchor.Id` remains the canonical full Unid identifier in all modes.
- Shared `DocxSessionOps.ProjectAnchor` propagates to WASM bridge + Python stdio (`project_anchor`).
- npm wrapper: `session.projectAnchor(anchorId, depth?)`, `ProjectionDepth` / `AnchorIdRendering` const enums.

## Test plan
- [x] `MD050` — `FullUnid` is default + existing behavior.
- [x] `MD051` — `Abbreviated` produces shorter ids with 4-char floor.
- [x] `MD052` — `Abbreviated` AnchorIndex has dual keys.
- [x] `MD053` — `Sequential` produces `1, 2, 3, …` per (kind, scope) bucket.
- [x] `MD054` — `Sequential` AnchorIndex has dual keys.
- [x] `DS295` — `ProjectAnchor` on paragraph returns just that paragraph.
- [x] `DS296` — `ProjectAnchor` on heading returns its section.
- [x] `DS297` — `ProjectAnchor` on last heading extends to end of parent.
- [x] `DS298` — `ProjectAnchor` on unknown anchor throws.
- [x] Playwright `project-anchor-and-rendering.spec.ts` — 5 assertion areas.
- [x] Full `dotnet test` passes.
- [x] `npm run build && npm test` passes.
- [x] Release build clean (modulo pre-existing #173).

## Note for reviewers
- `Sequential` mode ids are intentionally NOT stable across `Project()` calls. The reset-on-each-projection contract is documented and tested.
- `Anchor.Id` semantics unchanged: always the full Unid form, regardless of rendering mode. Only emitted tokens and alias keys change.
EOF
)" 2>&1 | tail -5
```

Expected: PR URL printed.

---

## Self-Review

**Spec coverage** (issue #167):
- §1 (`ProjectAnchor` + `ProjectionDepth`): Tasks 6-7. ✓
- §2 (`AnchorIdRendering` Abbreviated): Tasks 2-3. ✓
- §2 (Sequential): Tasks 4-5. ✓
- §3 (dual-keyed AnchorIndex): Tasks 3 + 5. ✓
- Acceptance criterion — "Token-counting test: `Abbreviated` produces ≥ 5% shorter output": not explicitly tested but `MD051` enforces `< 32` chars per token unid, which guarantees compression. A perf-style assertion could be added if desired.
- Acceptance criterion — "Anchor lookups work via both forms": `MD052` + `MD054`. ✓
- Cross-stack propagation: Tasks 8-9. ✓

**Placeholder scan:** No `TODO`, no `Similar to Task N`. All steps have real code.

**Type consistency:**
- `AnchorIdRendering` enum: FullUnid=0, Abbreviated=1, Sequential=2 (.NET + npm).
- `ProjectionDepth` enum: SelfOnly=0, Subtree=1, SubtreeAndFollowingSiblings=2.
- `Anchor.Id` semantics unchanged.

One spec deviation to flag:
- The `TryResolveToUnid` helper in `ProjectAnchor` does a roundtrip lookup through the AnchorIndex to translate rendered ids back to full Unids. For the default `FullUnid` case this is no-op (the rendered id IS the full Unid). For Abbreviated/Sequential it uses the dual-key index. This is the simplest correct approach but assumes the `AnchorIndex` is dual-keyed — which holds because the same `Settings.AnchorIdRendering` value is used for both the initial `Project()` and any subsequent `ProjectAnchor()` calls on the same session.

---

**Plan complete and saved to `docs/superpowers/plans/2026-05-25-issue-167-project-anchor-and-anchor-id-rendering.md`.**

Two execution options:

1. **Subagent-Driven (recommended)** — same flow as prior issues.
2. **Inline Execution** — execute tasks here with checkpoints.
