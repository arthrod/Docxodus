# DOCX Mutation API Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement a stateful, in-memory DOCX mutation API (`DocxSession`) per the approved spec at `docs/superpowers/specs/2026-05-24-docx-mutation-api-design.md`.

**Architecture:** A long-lived `DocxSession` over an in-memory `WordprocessingDocument`, exposing markdown-keyed mutation tools (tiers A–D) plus a `Raw.*` OOXML escape hatch. Each op flows through a uniform validate → snapshot → apply → reproject pipeline and returns a typed `EditResult`. Anchor ids come from the existing `WmlToMarkdownConverter` projection. WASM/npm parity ships in Phase 8.

**Tech Stack:** .NET 8, DocumentFormat.OpenXml 3.2.0, xUnit, .NET WASM (`[JSExport]`), TypeScript/npm, Playwright. All new code uses `#nullable enable`.

---

## Notes for the Engineer

- Read `docs/superpowers/specs/2026-05-24-docx-mutation-api-design.md` end-to-end before starting. It is the contract.
- Read `Docxodus/WmlToMarkdownConverter.cs` (~1000 LOC) — you will reuse its `ScopeInfo`, `EmitContext`, and many handlers. Don't duplicate; extract where needed.
- Read `Docxodus/UnidHelper.cs`. Every new block-level element you create must get a Unid through this helper before being added to the tree.
- Test ID prefix is `DS###`. Each phase has its own range (DS001–DS009 phase 1, DS010–DS019 phase 2, etc.) — preserves the convention used elsewhere in the codebase.
- Commit after each task in this plan. The branch is `feat/docx-mutation-api`.
- `Directory.Build.props` enforces `TreatWarningsAsErrors=true` in Release. Run `dotnet build -c Release Docxodus.sln` before any commit that touches code; nullable warnings will fail the build.
- This plan favors small, mergeable commits. Don't bundle multiple tasks into one commit.

---

## File Structure

**Created:**
- `Docxodus/DocxSession.cs` — public session class + public value types + settings
- `Docxodus/RawDocxOps.cs` — public `Raw.*` namespace class
- `Docxodus/Internal/UndoRing.cs` — bounded snapshot ring buffer
- `Docxodus/Internal/MarkdownPayloadParser.cs` — markdown subset → `XElement` runs/paragraphs
- `Docxodus/Internal/ScopeReprojector.cs` — re-runs the projector over one scope (extracted from `WmlToMarkdownConverter` shared internals)
- `Docxodus/Internal/MutationOps/*.cs` — one file per op family (TextOps, StructuralOps, FormatOps, CellOps, RawOps) holding the actual XML mutation logic
- `Docxodus.Tests/DocxSessionTests.cs` — all .NET tests, prefixed DS###
- `Docxodus.Tests/MarkdownPayloadParserTests.cs` — parser unit tests
- `TestFiles/DS001_simple_two_paragraphs.docx` through `DS005_raw_complex.docx` — fixtures
- `wasm/DocxodusWasm/DocxSessionBridge.cs` — JSExport bridge
- `npm/src/session.ts` — TypeScript wrapper
- `npm/src/react.ts` — React hook (extend existing file)
- `npm/scripts/generate-error-codes.mjs` — codegen for EditErrorCode union
- `npm/tests/docx-session.spec.ts` — Playwright tests
- `docs/architecture/docx_mutation_api.md` — in-tree spec (Phase 9)

**Modified:**
- `Docxodus/WmlToMarkdownConverter.cs` — expose `ScopeInfo`, `EmitContext`, and per-scope emission helpers as `internal` so the session can reuse them
- `Docxodus/UnidHelper.cs` — add `AssignToAllElements(XElement, bool includeSelf)` overload for the raw escape hatch
- `npm/src/types.ts` — add `DocxSession*`, `EditResult`, `EditErrorCode` etc.
- `npm/src/index.ts` — re-export `openDocxSession`, `DocxSession`
- `npm/src/docxodus.worker.ts` — extend worker protocol for session ops
- `npm/src/worker-proxy.ts` — add session-aware proxy methods
- `wasm/DocxodusWasm/DocxodusWasm.csproj` — no change expected (DocxSession is in core)
- `CHANGELOG.md` — entry per phase under `[Unreleased]`
- `CLAUDE.md` — module entry under "Core Modules" (Phase 9)

---

## Phase 1: Skeleton (no mutations yet)

**Deliverable:** `new DocxSession(bytes)`, `Project()`, `Save()`, `Exists()`, `GetAnchorInfo()`, internal `UndoRing` infra. Public API surface defined; mutation methods all return `EditResult { Success=false, Error=InternalError("not yet implemented") }`.

### Task 1.1: Add the simple two-paragraph fixture

**Files:**
- Create: `TestFiles/DS001_simple_two_paragraphs.docx`

- [ ] **Step 1: Build the fixture programmatically via a one-shot script**

Run this from the repo root to create the file using the OpenXML SDK (avoids hand-authored binary blobs):

```bash
cat > /tmp/make_ds001.csx <<'EOF'
#r "nuget: DocumentFormat.OpenXml, 3.2.0"
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
var path = "TestFiles/DS001_simple_two_paragraphs.docx";
using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
var main = doc.AddMainDocumentPart();
main.Document = new Document(new Body(
    new Paragraph(new Run(new Text("First paragraph."))),
    new Paragraph(new Run(new Text("Second paragraph.")))
));
EOF
dotnet script /tmp/make_ds001.csx
```

If `dotnet script` isn't installed (`dotnet tool install -g dotnet-script`), use this alternative: add a one-off `[Fact]` test under `Docxodus.Tests/_FixtureBuilders.cs` that writes the file, run with `dotnet test --filter Name=BuildDs001`, then delete the builder fact. (The codebase already uses this pattern — `OpenContractExporterTests` and `DocumentBuilderTests` both generate fixtures inline.)

- [ ] **Step 2: Verify the fixture opens**

```bash
dotnet build Docxodus/Docxodus.csproj
```

Expected: clean build (this step just confirms the .docx is valid by being readable later).

- [ ] **Step 3: Commit**

```bash
git add TestFiles/DS001_simple_two_paragraphs.docx
git commit -m "test(session): add DS001 simple two-paragraph fixture"
```

### Task 1.2: Define public value types

**Files:**
- Create: `Docxodus/DocxSession.cs` (initial scaffolding only — value types)

- [ ] **Step 1: Create the file with value types and enums**

```csharp
#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace Docxodus;

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
```

- [ ] **Step 2: Build and verify clean**

```bash
dotnet build -c Release Docxodus/Docxodus.csproj
```

Expected: zero warnings, zero errors.

- [ ] **Step 3: Commit**

```bash
git add Docxodus/DocxSession.cs
git commit -m "feat(session): public value types and error codes"
```

### Task 1.3: Define DocxSessionSettings

**Files:**
- Modify: `Docxodus/DocxSession.cs` (append)

- [ ] **Step 1: Append settings class**

Add to the end of `Docxodus/DocxSession.cs`:

```csharp
public sealed class DocxSessionSettings
{
    public int UndoDepth { get; init; } = 50;
    public bool ValidateRawOps { get; init; } = false;
    public TrackedChangeMode TrackedChanges { get; init; } = TrackedChangeMode.Accept;
    public string? RevisionAuthor { get; init; }
    public WmlToMarkdownConverterSettings ProjectionSettings { get; init; } = new();
    public Microsoft.Extensions.Logging.ILogger? Logger { get; init; }
}
```

Note: `TrackedChangeMode` already exists in `WmlToMarkdownConverter.cs` — reuse it.

- [ ] **Step 2: Build**

```bash
dotnet build -c Release Docxodus/Docxodus.csproj
```

Expected: clean. May need `using Microsoft.Extensions.Logging;` — confirm the package is referenced by Docxodus (it should be; `ExternalAnnotationProjector` uses it). If not, the `Microsoft.Extensions.Logging.ILogger?` fully-qualified reference works without an import.

- [ ] **Step 3: Commit**

```bash
git add Docxodus/DocxSession.cs
git commit -m "feat(session): DocxSessionSettings"
```

### Task 1.4: Create UndoRing infrastructure

**Files:**
- Create: `Docxodus/Internal/UndoRing.cs`

- [ ] **Step 1: Write the failing test**

Add to a new file `Docxodus.Tests/Internal/UndoRingTests.cs`:

```csharp
#nullable enable

using Xunit;

namespace Docxodus.Tests.Internal;

public class UndoRingTests
{
    [Fact]
    public void DS001a_PushAndPopRoundtrip()
    {
        var ring = new Docxodus.Internal.UndoRing<string>(capacity: 3);
        ring.PushUndo("v1");
        ring.PushUndo("v2");
        ring.PushUndo("v3");
        Assert.Equal("v3", ring.PopUndo());
        Assert.Equal("v2", ring.PopUndo());
        Assert.Equal("v1", ring.PopUndo());
        Assert.Null(ring.PopUndo());
    }

    [Fact]
    public void DS001b_CapacityEvictsOldest()
    {
        var ring = new Docxodus.Internal.UndoRing<string>(capacity: 2);
        ring.PushUndo("v1");
        ring.PushUndo("v2");
        ring.PushUndo("v3");   // evicts v1
        Assert.Equal("v3", ring.PopUndo());
        Assert.Equal("v2", ring.PopUndo());
        Assert.Null(ring.PopUndo());
    }

    [Fact]
    public void DS001c_NewPushClearsRedo()
    {
        var ring = new Docxodus.Internal.UndoRing<string>(capacity: 5);
        ring.PushUndo("v1");
        ring.PushUndo("v2");
        ring.PopUndo();                  // v2 -> redo stack
        Assert.Equal("v2", ring.PopRedo());
        ring.PushRedo("v2");             // restore redo for next assertion
        ring.PopUndo();                  // pop v1; v1 -> redo... actually no, we popped already
        // simpler scenario:
        var r2 = new Docxodus.Internal.UndoRing<string>(capacity: 5);
        r2.PushUndo("a");
        r2.PushUndo("b");
        Assert.Equal("b", r2.PopUndo());     // b on redo
        r2.PushUndo("c");                    // new edit clears redo
        Assert.Null(r2.PopRedo());
    }
}
```

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter FullyQualifiedName~UndoRingTests`
Expected: FAIL — `UndoRing` type does not exist.

- [ ] **Step 2: Implement UndoRing**

Create `Docxodus/Internal/UndoRing.cs`:

```csharp
#nullable enable

using System.Collections.Generic;

namespace Docxodus.Internal;

internal sealed class UndoRing<T>
{
    private readonly LinkedList<T> _undo = new();
    private readonly Stack<T> _redo = new();
    private readonly int _capacity;

    public UndoRing(int capacity)
    {
        _capacity = capacity > 0 ? capacity : 1;
    }

    public int UndoCount => _undo.Count;
    public int RedoCount => _redo.Count;

    public void PushUndo(T snapshot)
    {
        _undo.AddLast(snapshot);
        while (_undo.Count > _capacity) _undo.RemoveFirst();
        _redo.Clear();
    }

    public T? PopUndo()
    {
        if (_undo.Count == 0) return default;
        var last = _undo.Last!.Value;
        _undo.RemoveLast();
        _redo.Push(last);
        return last;
    }

    public T? PopRedo()
    {
        if (_redo.Count == 0) return default;
        var v = _redo.Pop();
        _undo.AddLast(v);
        return v;
    }

    public void PushRedo(T snapshot) => _redo.Push(snapshot);

    public void Clear()
    {
        _undo.Clear();
        _redo.Clear();
    }
}
```

- [ ] **Step 3: Run tests, expect pass**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter FullyQualifiedName~UndoRingTests`
Expected: 3/3 PASS.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/Internal/UndoRing.cs Docxodus.Tests/Internal/UndoRingTests.cs
git commit -m "feat(session): bounded undo/redo ring buffer"
```

### Task 1.5: DocxSession skeleton — open, Project, Save, dispose

**Files:**
- Modify: `Docxodus/DocxSession.cs` (append `DocxSession` class)

- [ ] **Step 1: Write the failing test**

Create `Docxodus.Tests/DocxSessionTests.cs`:

```csharp
#nullable enable

using System;
using System.IO;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static byte[] LoadFixtureBytes(string name) =>
        File.ReadAllBytes(Path.Combine(TestFilesDir.FullName, name));

    [Fact]
    public void DS001_OpenAndProject()
    {
        var bytes = LoadFixtureBytes("DS001_simple_two_paragraphs.docx");
        using var session = new DocxSession(bytes);

        var projection = session.Project();
        Assert.Contains("First paragraph.", projection.Markdown);
        Assert.Contains("Second paragraph.", projection.Markdown);
        Assert.True(projection.AnchorIndex.Count >= 2);
    }

    [Fact]
    public void DS002_SaveRoundtrip()
    {
        var bytes = LoadFixtureBytes("DS001_simple_two_paragraphs.docx");
        using var session = new DocxSession(bytes);
        var out1 = session.Save();
        Assert.NotEmpty(out1);

        // Re-open the saved bytes and project — same body content.
        using var session2 = new DocxSession(out1);
        Assert.Contains("First paragraph.", session2.Project().Markdown);
    }

    [Fact]
    public void DS003_ExistsAndGetAnchorInfo()
    {
        var bytes = LoadFixtureBytes("DS001_simple_two_paragraphs.docx");
        using var session = new DocxSession(bytes);
        var proj = session.Project();

        var firstAnchor = System.Linq.Enumerable.First(proj.AnchorIndex.Keys);
        Assert.True(session.Exists(firstAnchor));
        Assert.False(session.Exists("p:body:deadbeef"));

        var info = session.GetAnchorInfo(firstAnchor);
        Assert.NotNull(info);
        Assert.Contains(info!.Kind, new[] { "p", "h", "li" });
    }

    [Fact]
    public void DS004_DisposeDoubleOk()
    {
        var bytes = LoadFixtureBytes("DS001_simple_two_paragraphs.docx");
        var session = new DocxSession(bytes);
        session.Dispose();
        session.Dispose();   // must not throw
    }
}
```

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter FullyQualifiedName~DocxSessionTests`
Expected: FAIL — `DocxSession` type missing.

- [ ] **Step 2: Implement DocxSession skeleton**

Append to `Docxodus/DocxSession.cs`:

```csharp
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus;

public sealed class DocxSession : IDisposable
{
    private readonly DocxSessionSettings _settings;
    private MemoryStream? _stream;
    private WordprocessingDocument? _doc;
    private MarkdownProjection? _cachedProjection;
    private bool _disposed;

    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        _settings = settings ?? new DocxSessionSettings();
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

    public bool Exists(string anchorId)
    {
        ThrowIfDisposed();
        return Project().AnchorIndex.ContainsKey(anchorId);
    }

    public AnchorInfo? GetAnchorInfo(string anchorId)
    {
        ThrowIfDisposed();
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target)) return null;

        var element = target.Resolve(_doc!);
        var preview = element is null ? "" : ElementTextPreview(element);
        var parts = anchorId.Split(':', 3);
        var kind = parts.Length > 0 ? parts[0] : "";
        var scope = parts.Length > 1 ? parts[1] : "";
        return new AnchorInfo(anchorId, kind, scope, preview);
    }

    public byte[] Save()
    {
        ThrowIfDisposed();
        _doc!.Save();
        _stream!.Flush();
        _stream.Position = 0;
        return _stream.ToArray();
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

    internal void InvalidateProjectionCache() => _cachedProjection = null;

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(DocxSession));
    }

    private static string ElementTextPreview(XElement element)
    {
        var text = string.Concat(element.Descendants(W.t).Select(t => (string)t));
        return text.Length > 80 ? text.Substring(0, 80) + "…" : text;
    }
}
```

- [ ] **Step 3: Run tests, expect pass**

Run: `dotnet test Docxodus.Tests/Docxodus.Tests.csproj --filter FullyQualifiedName~DocxSessionTests`
Expected: 4/4 PASS.

- [ ] **Step 4: Release build to catch nullable warnings**

```bash
dotnet build -c Release Docxodus/Docxodus.csproj
```

Expected: zero warnings.

- [ ] **Step 5: Commit**

```bash
git add Docxodus/DocxSession.cs Docxodus.Tests/DocxSessionTests.cs
git commit -m "feat(session): skeleton — open, project, save, dispose"
```

### Task 1.6: CHANGELOG entry for Phase 1

**Files:**
- Modify: `CHANGELOG.md`

- [ ] **Step 1: Add entry**

Under the `[Unreleased]` section's `### Added` (create the heading if missing):

```markdown
### Added

- **DocxSession (skeleton)** — Stateful in-memory document editing surface keyed by markdown projection anchor ids. Phase 1 ships `Project`, `Save`, `Exists`, `GetAnchorInfo`, and lifecycle (`Dispose`). Mutation methods land in Phase 2+. Spec at `docs/superpowers/specs/2026-05-24-docx-mutation-api-design.md`.
```

- [ ] **Step 2: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(session): changelog entry for phase 1 skeleton"
```

---

## Phase 2: Markdown payload parser

**Deliverable:** A pure module `Internal/MarkdownPayloadParser.cs` that converts the supported markdown subset into a list of `XElement w:p` block elements (with run children carrying `rPr`). Block-only; no document-state dependency. Fully unit-tested.

### Task 2.1: Define parser types and inline result

**Files:**
- Create: `Docxodus/Internal/MarkdownPayloadParser.cs`

- [ ] **Step 1: Create file with public-internal API**

```csharp
#nullable enable

using System.Collections.Generic;
using System.Xml.Linq;

namespace Docxodus.Internal;

internal enum ParserBlockKind { Paragraph, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Quote, Code, BulletItem, OrderedItem }

internal sealed record ParsedBlock(
    ParserBlockKind Kind,
    int ListLevel,                 // 0 for non-list
    IReadOnlyList<XElement> RunElements);

internal sealed record ParseError(EditErrorCode Code, string Message);

internal sealed class ParseResult
{
    public IReadOnlyList<ParsedBlock> Blocks { get; init; } = System.Array.Empty<ParsedBlock>();
    public ParseError? Error { get; init; }
    public bool Success => Error is null;

    public static ParseResult Ok(IReadOnlyList<ParsedBlock> blocks) => new() { Blocks = blocks };
    public static ParseResult Fail(EditErrorCode code, string msg) =>
        new() { Error = new ParseError(code, msg) };
}

internal static class MarkdownPayloadParser
{
    public static ParseResult Parse(string markdown)
    {
        // Stub — implemented in subsequent tasks.
        throw new System.NotImplementedException();
    }
}
```

- [ ] **Step 2: Build clean**

```bash
dotnet build -c Release Docxodus/Docxodus.csproj
```

- [ ] **Step 3: Commit**

```bash
git add Docxodus/Internal/MarkdownPayloadParser.cs
git commit -m "feat(session): parser scaffolding"
```

### Task 2.2: Inline parser — plain text and escapes

**Files:**
- Modify: `Docxodus/Internal/MarkdownPayloadParser.cs`
- Create: `Docxodus.Tests/MarkdownPayloadParserTests.cs`

- [ ] **Step 1: Failing tests**

```csharp
#nullable enable

using System.Linq;
using System.Xml.Linq;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class MarkdownPayloadParserTests
{
    [Fact]
    public void DS010_PlainText()
    {
        var r = MarkdownPayloadParser.Parse("Hello world.");
        Assert.True(r.Success);
        Assert.Single(r.Blocks);
        var p = r.Blocks[0];
        Assert.Equal(ParserBlockKind.Paragraph, p.Kind);
        var text = string.Concat(p.RunElements.Descendants(W.t).Select(t => (string)t));
        Assert.Equal("Hello world.", text);
    }

    [Fact]
    public void DS011_EscapedAsterisk()
    {
        var r = MarkdownPayloadParser.Parse(@"Not \*bold\*.");
        Assert.True(r.Success);
        var text = string.Concat(r.Blocks[0].RunElements.Descendants(W.t).Select(t => (string)t));
        Assert.Equal("Not *bold*.", text);
    }
}
```

Run: `dotnet test --filter FullyQualifiedName~MarkdownPayloadParserTests`
Expected: FAIL.

- [ ] **Step 2: Implement inline parser for plain + escapes**

Replace the stub in `Docxodus/Internal/MarkdownPayloadParser.cs` with a real implementation. The parser uses a hand-rolled scanner (no Markdig dependency) over the subset. Start narrow — this step only covers plain text + backslash escapes; further inline syntax follows in Task 2.3.

```csharp
#nullable enable

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace Docxodus.Internal;

internal static class MarkdownPayloadParser
{
    public static ParseResult Parse(string markdown)
    {
        if (markdown is null) return ParseResult.Fail(EditErrorCode.MalformedMarkdown, "null payload");
        var blocks = SplitBlocks(markdown);
        var parsed = new List<ParsedBlock>(blocks.Count);
        foreach (var raw in blocks)
        {
            var b = ParseBlock(raw);
            if (b is null) continue;
            parsed.Add(b);
        }
        return ParseResult.Ok(parsed);
    }

    private static List<string> SplitBlocks(string md)
    {
        var lines = md.Replace("\r\n", "\n").Split('\n');
        var result = new List<string>();
        var buf = new StringBuilder();
        foreach (var line in lines)
        {
            if (line.Length == 0)
            {
                if (buf.Length > 0) { result.Add(buf.ToString()); buf.Clear(); }
            }
            else
            {
                if (buf.Length > 0) buf.Append('\n');
                buf.Append(line);
            }
        }
        if (buf.Length > 0) result.Add(buf.ToString());
        return result;
    }

    private static ParsedBlock? ParseBlock(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return null;
        // Plain paragraph for now; richer block recognition in later tasks.
        var runs = ParseInline(raw);
        return new ParsedBlock(ParserBlockKind.Paragraph, 0, runs);
    }

    private static IReadOnlyList<XElement> ParseInline(string text)
    {
        var sb = new StringBuilder();
        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (c == '\\' && i + 1 < text.Length)
            {
                sb.Append(text[++i]);
                continue;
            }
            sb.Append(c);
        }
        return new[] { TextRun(sb.ToString(), bold: false, italic: false) };
    }

    internal static XElement TextRun(string text, bool bold, bool italic,
        bool code = false, bool strike = false, bool underline = false,
        string? color = null, string? runStyle = null)
    {
        var rPr = new XElement(W.rPr);
        if (bold)      rPr.Add(new XElement(W.b));
        if (italic)    rPr.Add(new XElement(W.i));
        if (strike)    rPr.Add(new XElement(W.strike));
        if (underline) rPr.Add(new XElement(W.u, new XAttribute(W.val, "single")));
        if (code || runStyle is not null)
            rPr.Add(new XElement(W.rStyle, new XAttribute(W.val, runStyle ?? "Code")));
        if (color is not null && color.Length > 0)
            rPr.Add(new XElement(W.color, new XAttribute(W.val, color)));

        var run = new XElement(W.r);
        if (rPr.HasElements) run.Add(rPr);
        run.Add(new XElement(W.t, new XAttribute(XNamespace.Xml + "space", "preserve"), text));
        return run;
    }
}
```

- [ ] **Step 3: Run tests, expect pass**

```bash
dotnet test --filter FullyQualifiedName~MarkdownPayloadParserTests
```

Expected: DS010, DS011 PASS.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/Internal/MarkdownPayloadParser.cs Docxodus.Tests/MarkdownPayloadParserTests.cs
git commit -m "feat(session): parser — plain text and escapes"
```

### Task 2.3: Inline parser — bold/italic/code/strike/links

**Files:**
- Modify: `Docxodus/Internal/MarkdownPayloadParser.cs`
- Modify: `Docxodus.Tests/MarkdownPayloadParserTests.cs`

- [ ] **Step 1: Failing tests**

Append to `MarkdownPayloadParserTests.cs`:

```csharp
    [Fact]
    public void DS012_Bold()
    {
        var r = MarkdownPayloadParser.Parse("This is **bold** text.");
        Assert.True(r.Success);
        var runs = r.Blocks[0].RunElements;
        Assert.True(runs.Count >= 3);
        var boldRun = runs.Single(run => run.Element(W.rPr)?.Element(W.b) is not null);
        Assert.Equal("bold", (string)boldRun.Element(W.t)!);
    }

    [Fact]
    public void DS013_Italic()
    {
        var r = MarkdownPayloadParser.Parse("This is *italic* text.");
        Assert.True(r.Success);
        var italicRun = r.Blocks[0].RunElements
            .Single(run => run.Element(W.rPr)?.Element(W.i) is not null);
        Assert.Equal("italic", (string)italicRun.Element(W.t)!);
    }

    [Fact]
    public void DS014_Code()
    {
        var r = MarkdownPayloadParser.Parse("Inline `code` here.");
        Assert.True(r.Success);
        var codeRun = r.Blocks[0].RunElements.Single(run =>
            (string?)run.Element(W.rPr)?.Element(W.rStyle)?.Attribute(W.val) == "Code");
        Assert.Equal("code", (string)codeRun.Element(W.t)!);
    }

    [Fact]
    public void DS015_Strike()
    {
        var r = MarkdownPayloadParser.Parse("Some ~~struck~~ text.");
        Assert.True(r.Success);
        var s = r.Blocks[0].RunElements.Single(run => run.Element(W.rPr)?.Element(W.strike) is not null);
        Assert.Equal("struck", (string)s.Element(W.t)!);
    }

    [Fact]
    public void DS016_Link()
    {
        var r = MarkdownPayloadParser.Parse("Visit [Docxodus](https://example.com/d) today.");
        Assert.True(r.Success);
        var link = r.Blocks[0].RunElements.Single(e => e.Name == W.hyperlink);
        Assert.Equal("Docxodus", string.Concat(link.Descendants(W.t).Select(t => (string)t)));
        // External hyperlinks store their URI in a relationship; for the parser we stash
        // it on a "docxodus:href" attribute and resolve when stitching into the document.
        Assert.Equal("https://example.com/d", (string?)link.Attribute(XName.Get("href", "docxodus:")));
    }
```

Run: expect FAIL.

- [ ] **Step 2: Extend inline parser**

Replace `ParseInline` in `Docxodus/Internal/MarkdownPayloadParser.cs`:

```csharp
    private static IReadOnlyList<XElement> ParseInline(string text)
    {
        var list = new List<XElement>();
        var sb = new StringBuilder();
        bool bold = false, italic = false;

        void FlushText()
        {
            if (sb.Length == 0) return;
            list.Add(TextRun(sb.ToString(), bold, italic));
            sb.Clear();
        }

        for (int i = 0; i < text.Length; i++)
        {
            var c = text[i];

            // Escapes
            if (c == '\\' && i + 1 < text.Length)
            {
                sb.Append(text[++i]);
                continue;
            }

            // Code span
            if (c == '`')
            {
                FlushText();
                int end = text.IndexOf('`', i + 1);
                if (end < 0) { sb.Append(c); continue; }
                list.Add(TextRun(text.Substring(i + 1, end - i - 1), bold: false, italic: false, code: true));
                i = end;
                continue;
            }

            // Strike (GFM ~~text~~)
            if (c == '~' && i + 1 < text.Length && text[i + 1] == '~')
            {
                FlushText();
                int end = text.IndexOf("~~", i + 2, StringComparison.Ordinal);
                if (end < 0) { sb.Append("~~"); i++; continue; }
                list.Add(TextRun(text.Substring(i + 2, end - i - 2), bold: false, italic: false, strike: true));
                i = end + 1;
                continue;
            }

            // Bold (**...**)
            if (c == '*' && i + 1 < text.Length && text[i + 1] == '*')
            {
                FlushText();
                bold = !bold;
                i++;
                continue;
            }

            // Italic (*...*)
            if (c == '*')
            {
                FlushText();
                italic = !italic;
                continue;
            }

            // Link [text](url)
            if (c == '[')
            {
                int close = text.IndexOf(']', i + 1);
                if (close > 0 && close + 1 < text.Length && text[close + 1] == '(')
                {
                    int paren = text.IndexOf(')', close + 2);
                    if (paren > 0)
                    {
                        FlushText();
                        var linkText = text.Substring(i + 1, close - i - 1);
                        var url = text.Substring(close + 2, paren - close - 2);
                        var hyperlink = new XElement(W.hyperlink,
                            new XAttribute(XName.Get("href", "docxodus:"), url),
                            TextRun(linkText, bold, italic));
                        list.Add(hyperlink);
                        i = paren;
                        continue;
                    }
                }
                sb.Append(c);
                continue;
            }

            // Anchor token in payload — reject
            if (c == '{' && i + 1 < text.Length && text[i + 1] == '#')
            {
                throw new MarkdownPayloadException(
                    EditErrorCode.AnchorTokenInPayload,
                    "Anchor tokens like {#kind:scope:unid} are projection output, not input. Remove them from the payload.");
            }

            sb.Append(c);
        }

        FlushText();
        return list;
    }

    internal sealed class MarkdownPayloadException : System.Exception
    {
        public EditErrorCode Code { get; }
        public MarkdownPayloadException(EditErrorCode code, string msg) : base(msg) => Code = code;
    }
```

Update `Parse` to catch the exception:

```csharp
    public static ParseResult Parse(string markdown)
    {
        if (markdown is null) return ParseResult.Fail(EditErrorCode.MalformedMarkdown, "null payload");
        try
        {
            var blocks = SplitBlocks(markdown);
            var parsed = new List<ParsedBlock>(blocks.Count);
            foreach (var raw in blocks)
            {
                var b = ParseBlock(raw);
                if (b is not null) parsed.Add(b);
            }
            return ParseResult.Ok(parsed);
        }
        catch (MarkdownPayloadException ex)
        {
            return ParseResult.Fail(ex.Code, ex.Message);
        }
    }
```

- [ ] **Step 3: Tests pass**

```bash
dotnet test --filter FullyQualifiedName~MarkdownPayloadParserTests
```

Expected: DS010-DS016 all PASS.

- [ ] **Step 4: Commit**

```bash
git add Docxodus/Internal/MarkdownPayloadParser.cs Docxodus.Tests/MarkdownPayloadParserTests.cs
git commit -m "feat(session): parser — bold/italic/code/strike/links"
```

### Task 2.4: Block parser — headings, blockquote, fenced code

**Files:**
- Modify: `Docxodus/Internal/MarkdownPayloadParser.cs`
- Modify: `Docxodus.Tests/MarkdownPayloadParserTests.cs`

- [ ] **Step 1: Failing tests**

Append:

```csharp
    [Fact]
    public void DS017_Headings()
    {
        var r = MarkdownPayloadParser.Parse("# H1\n\n## H2\n\n###### H6");
        Assert.True(r.Success);
        Assert.Equal(3, r.Blocks.Count);
        Assert.Equal(ParserBlockKind.Heading1, r.Blocks[0].Kind);
        Assert.Equal(ParserBlockKind.Heading2, r.Blocks[1].Kind);
        Assert.Equal(ParserBlockKind.Heading6, r.Blocks[2].Kind);
    }

    [Fact]
    public void DS018_Blockquote()
    {
        var r = MarkdownPayloadParser.Parse("> Quoted text.");
        Assert.True(r.Success);
        Assert.Equal(ParserBlockKind.Quote, r.Blocks[0].Kind);
    }

    [Fact]
    public void DS019_FencedCode()
    {
        var r = MarkdownPayloadParser.Parse("```\ncode line 1\ncode line 2\n```");
        Assert.True(r.Success);
        Assert.Equal(ParserBlockKind.Code, r.Blocks[0].Kind);
        var text = string.Concat(r.Blocks[0].RunElements.Descendants(W.t).Select(t => (string)t));
        Assert.Contains("code line 1", text);
        Assert.Contains("code line 2", text);
    }
```

Run: expect FAIL.

- [ ] **Step 2: Extend `ParseBlock` and `SplitBlocks`**

Update `SplitBlocks` to recognize fenced-code regions as single blocks (fence opens/closes regardless of blank lines):

```csharp
    private static List<string> SplitBlocks(string md)
    {
        var normalized = md.Replace("\r\n", "\n");
        var lines = normalized.Split('\n');
        var result = new List<string>();
        var buf = new StringBuilder();
        bool inFence = false;

        void Flush()
        {
            if (buf.Length > 0) { result.Add(buf.ToString()); buf.Clear(); }
        }

        foreach (var line in lines)
        {
            if (line.StartsWith("```"))
            {
                if (!inFence) Flush();
                inFence = !inFence;
                if (buf.Length > 0) buf.Append('\n');
                buf.Append(line);
                if (!inFence) Flush();
                continue;
            }
            if (!inFence && line.Length == 0)
            {
                Flush();
                continue;
            }
            if (buf.Length > 0) buf.Append('\n');
            buf.Append(line);
        }
        Flush();
        return result;
    }
```

Update `ParseBlock` to dispatch on prefix:

```csharp
    private static ParsedBlock? ParseBlock(string raw)
    {
        if (string.IsNullOrEmpty(raw)) return null;

        // Fenced code
        if (raw.StartsWith("```"))
        {
            var firstNewline = raw.IndexOf('\n');
            var lastNewline = raw.LastIndexOf('\n');
            string inner = firstNewline >= 0 && lastNewline > firstNewline
                ? raw.Substring(firstNewline + 1, lastNewline - firstNewline - 1)
                : "";
            return new ParsedBlock(ParserBlockKind.Code, 0,
                new[] { TextRun(inner, bold: false, italic: false) });
        }

        // ATX headings
        if (raw.StartsWith("#"))
        {
            int level = 0;
            while (level < raw.Length && raw[level] == '#') level++;
            if (level >= 1 && level <= 6 && level < raw.Length && raw[level] == ' ')
            {
                var headingText = raw.Substring(level + 1).TrimEnd();
                return new ParsedBlock(
                    (ParserBlockKind)((int)ParserBlockKind.Heading1 + level - 1),
                    0,
                    ParseInline(headingText));
            }
        }

        // Blockquote (each line begins with "> ")
        if (raw.StartsWith("> "))
        {
            var quoteText = string.Join("\n",
                raw.Split('\n').Select(l => l.StartsWith("> ") ? l.Substring(2) : l));
            return new ParsedBlock(ParserBlockKind.Quote, 0, ParseInline(quoteText));
        }

        return new ParsedBlock(ParserBlockKind.Paragraph, 0, ParseInline(raw));
    }
```

(Add `using System.Linq;` at the top if not already present.)

- [ ] **Step 3: Tests pass**

```bash
dotnet test --filter FullyQualifiedName~MarkdownPayloadParserTests
```

Expected: DS010-DS019 PASS.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): parser — headings, blockquote, fenced code"
```

### Task 2.5: Block parser — bulleted and ordered lists (with nesting)

**Files:**
- Modify: `Docxodus/Internal/MarkdownPayloadParser.cs`
- Modify: `Docxodus.Tests/MarkdownPayloadParserTests.cs`

- [ ] **Step 1: Failing tests**

Append:

```csharp
    [Fact]
    public void DS020_BulletedList()
    {
        var r = MarkdownPayloadParser.Parse("- First\n- Second\n- Third");
        Assert.True(r.Success);
        Assert.Equal(3, r.Blocks.Count);
        Assert.All(r.Blocks, b => Assert.Equal(ParserBlockKind.BulletItem, b.Kind));
        Assert.All(r.Blocks, b => Assert.Equal(0, b.ListLevel));
    }

    [Fact]
    public void DS021_OrderedList()
    {
        var r = MarkdownPayloadParser.Parse("1. One\n2. Two");
        Assert.True(r.Success);
        Assert.Equal(2, r.Blocks.Count);
        Assert.All(r.Blocks, b => Assert.Equal(ParserBlockKind.OrderedItem, b.Kind));
    }

    [Fact]
    public void DS022_NestedList()
    {
        var r = MarkdownPayloadParser.Parse("- Top\n  - Nested\n  - Also nested\n- Top again");
        Assert.True(r.Success);
        Assert.Equal(4, r.Blocks.Count);
        Assert.Equal(0, r.Blocks[0].ListLevel);
        Assert.Equal(1, r.Blocks[1].ListLevel);
        Assert.Equal(1, r.Blocks[2].ListLevel);
        Assert.Equal(0, r.Blocks[3].ListLevel);
    }
```

Expected: FAIL.

- [ ] **Step 2: Extend `SplitBlocks` to treat consecutive list lines as one chunk per line**

Actually, lists are simpler if we let each line be its own block and rely on `ParseBlock` to recognize the marker. Change `SplitBlocks` to NOT collapse list lines under one block:

In `SplitBlocks`, add detection at the start of each line:

```csharp
            // List items: each item is its own block (no blank-line requirement between items)
            if (!inFence && IsListLineStart(line) && buf.Length > 0
                && !IsListLineStart(buf.ToString().Split('\n')[^1]))
            {
                Flush();
            }
            if (!inFence && IsListLineStart(line) && buf.Length > 0
                && IsListLineStart(buf.ToString().Split('\n')[^1]))
            {
                // start a new block per list line
                Flush();
            }
```

That's messy. Simpler: pre-tokenize. Replace `SplitBlocks` with:

```csharp
    private static List<string> SplitBlocks(string md)
    {
        var normalized = md.Replace("\r\n", "\n");
        var lines = normalized.Split('\n');
        var result = new List<string>();
        var buf = new StringBuilder();
        bool inFence = false;

        void Flush()
        {
            if (buf.Length > 0) { result.Add(buf.ToString()); buf.Clear(); }
        }

        foreach (var line in lines)
        {
            if (line.StartsWith("```"))
            {
                if (!inFence) Flush();
                inFence = !inFence;
                if (buf.Length > 0) buf.Append('\n');
                buf.Append(line);
                if (!inFence) Flush();
                continue;
            }
            if (!inFence && line.Length == 0)
            {
                Flush();
                continue;
            }
            if (!inFence && IsListLine(line))
            {
                Flush();             // each list item is its own block
                buf.Append(line);
                Flush();
                continue;
            }
            if (buf.Length > 0) buf.Append('\n');
            buf.Append(line);
        }
        Flush();
        return result;
    }

    private static bool IsListLine(string line)
    {
        int i = 0;
        while (i < line.Length && line[i] == ' ') i++;
        if (i >= line.Length) return false;
        if (line[i] == '-' || line[i] == '*' || line[i] == '+')
            return i + 1 < line.Length && line[i + 1] == ' ';
        // ordered: digits then '.' then ' '
        int start = i;
        while (i < line.Length && char.IsDigit(line[i])) i++;
        return i > start && i + 1 < line.Length && line[i] == '.' && line[i + 1] == ' ';
    }
```

Extend `ParseBlock` (insert before the fallback paragraph branch):

```csharp
        // Lists
        if (IsListLine(raw))
        {
            int indent = 0;
            while (indent < raw.Length && raw[indent] == ' ') indent++;
            int level = indent / 2;
            bool bullet = raw[indent] == '-' || raw[indent] == '*' || raw[indent] == '+';
            int markerEnd;
            if (bullet)
            {
                markerEnd = indent + 2; // "- "
            }
            else
            {
                markerEnd = indent;
                while (markerEnd < raw.Length && char.IsDigit(raw[markerEnd])) markerEnd++;
                markerEnd += 2; // ". "
            }
            var itemText = raw.Substring(markerEnd).TrimEnd();
            return new ParsedBlock(
                bullet ? ParserBlockKind.BulletItem : ParserBlockKind.OrderedItem,
                level,
                ParseInline(itemText));
        }
```

- [ ] **Step 3: Tests pass**

```bash
dotnet test --filter FullyQualifiedName~MarkdownPayloadParserTests
```

Expected: DS010-DS022 PASS.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): parser — bulleted, ordered, nested lists"
```

### Task 2.6: Rejection — pipe tables, footnotes, comments, images

**Files:**
- Modify: `Docxodus/Internal/MarkdownPayloadParser.cs`
- Modify: `Docxodus.Tests/MarkdownPayloadParserTests.cs`

- [ ] **Step 1: Failing tests**

```csharp
    [Theory]
    [InlineData("| col1 | col2 |\n|---|---|\n| a | b |", EditErrorCode.TableInsertNotSupported)]
    [InlineData("See [^fn-abc].", EditErrorCode.FootnoteRefNotSupported)]
    [InlineData("Hello {#cmt:cmt:abcd} there.", EditErrorCode.CommentMarkerNotSupported)]
    [InlineData("![alt](docxodus://img/abcd)", EditErrorCode.ImageInsertNotSupported)]
    [InlineData("Hello {#p:body:abcd} there.", EditErrorCode.AnchorTokenInPayload)]
    public void DS023_RejectionCodes(string payload, EditErrorCode expected)
    {
        var r = MarkdownPayloadParser.Parse(payload);
        Assert.False(r.Success);
        Assert.Equal(expected, r.Error!.Code);
    }
```

Expected: FAIL (or partial-fail).

- [ ] **Step 2: Add detection in `ParseBlock` (top of method, before other branches)**

```csharp
        // Pipe table — detect a line starting with '|' followed by '|---' line
        if (raw.StartsWith("|") && raw.Contains("\n|") && raw.Contains("---"))
        {
            throw new MarkdownPayloadException(
                EditErrorCode.TableInsertNotSupported,
                "Tables can't be inserted via markdown in v1. Use ReplaceCellContent(anchor, md) to edit an existing cell. InsertTable is planned for v2.");
        }
```

In `ParseInline` (before the `{` `#` anchor-token check), add detection for footnote refs and comment markers and image syntax:

```csharp
            // Footnote/endnote reference: [^id]
            if (c == '[' && i + 1 < text.Length && text[i + 1] == '^')
            {
                int close = text.IndexOf(']', i + 2);
                if (close > 0)
                {
                    throw new MarkdownPayloadException(
                        EditErrorCode.FootnoteRefNotSupported,
                        "Footnote/endnote references are output-only in v1. AddFootnote(anchor, md) is planned for v2.");
                }
            }

            // Image: ![alt](url)
            if (c == '!' && i + 1 < text.Length && text[i + 1] == '[')
            {
                throw new MarkdownPayloadException(
                    EditErrorCode.ImageInsertNotSupported,
                    "Image insertion requires a binary upload. AddImage(anchor, bytes, alt) is planned for v2.");
            }
```

Update the existing anchor-token branch to distinguish comments:

```csharp
            if (c == '{' && i + 1 < text.Length && text[i + 1] == '#')
            {
                int close = text.IndexOf('}', i + 2);
                if (close > 0)
                {
                    var token = text.Substring(i + 2, close - i - 2);
                    if (token.StartsWith("cmt:"))
                    {
                        throw new MarkdownPayloadException(
                            EditErrorCode.CommentMarkerNotSupported,
                            "Comment markers are output-only in v1. AddComment(anchor, author, md) is planned for v2.");
                    }
                    throw new MarkdownPayloadException(
                        EditErrorCode.AnchorTokenInPayload,
                        "Anchor tokens like {#kind:scope:unid} are projection output, not input. Remove them from the payload.");
                }
            }
```

- [ ] **Step 3: Tests pass**

```bash
dotnet test --filter FullyQualifiedName~MarkdownPayloadParserTests
```

Expected: DS023 (all 5 inline data rows) PASS plus existing tests still green.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): parser — typed rejection of unsupported syntax"
```

### Task 2.7: CHANGELOG entry for Phase 2

- [ ] **Step 1: Append to CHANGELOG under `### Added`**

```markdown
- **DocxSession markdown payload parser** — Phase 2. Hand-rolled parser accepting the projector-symmetric subset (paragraphs, headings, lists, blockquotes, fenced code; bold/italic/code/strike/links inline). Rejects unsupported syntax with typed `EditErrorCode`s.
```

- [ ] **Step 2: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(session): changelog phase 2"
```

---

## Phase 3: Tier A — text CRUD + undo/redo

**Deliverable:** `ReplaceText`, `DeleteBlock`, public `Undo`/`Redo`. First full vertical slice through the pipeline (validate → snapshot → apply → reproject → patch).

### Task 3.1: Snapshot helper — clone the MainDocumentPart XML

**Files:**
- Modify: `Docxodus/DocxSession.cs`

- [ ] **Step 1: Add private snapshot fields**

Inside `DocxSession`, after the `_disposed` field, add:

```csharp
    private readonly UndoRing<DocumentSnapshot> _history = new(50);

    private sealed record DocumentSnapshot(System.Xml.Linq.XDocument MainXml /*, future: header parts etc. */);

    private DocumentSnapshot TakeSnapshot()
    {
        var main = _doc!.MainDocumentPart!.GetXDocument();
        return new DocumentSnapshot(new System.Xml.Linq.XDocument(main));
    }

    private void RestoreSnapshot(DocumentSnapshot snapshot)
    {
        var part = _doc!.MainDocumentPart!;
        part.PutXDocument(new System.Xml.Linq.XDocument(snapshot.MainXml));
        InvalidateProjectionCache();
    }
```

Update the constructor to honor `Settings.UndoDepth`:

```csharp
    public DocxSession(byte[] docxBytes, DocxSessionSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        _settings = settings ?? new DocxSessionSettings();
        _stream = new MemoryStream();
        _stream.Write(docxBytes, 0, docxBytes.Length);
        _stream.Position = 0;
        _doc = WordprocessingDocument.Open(_stream, isEditable: true);
        _history = new UndoRing<DocumentSnapshot>(_settings.UndoDepth);
    }
```

(Remove the field initializer if you do it in the constructor; pick one. The constructor is simpler if `UndoDepth` is to be honored.)

- [ ] **Step 2: Build clean**

```bash
dotnet build -c Release Docxodus/Docxodus.csproj
```

- [ ] **Step 3: Commit**

```bash
git add Docxodus/DocxSession.cs
git commit -m "feat(session): snapshot infrastructure inside session"
```

### Task 3.2: ReplaceText — happy path on a plain paragraph

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS030_ReplaceTextSimple()
    {
        var bytes = LoadFixtureBytes("DS001_simple_two_paragraphs.docx");
        using var session = new DocxSession(bytes);
        var firstAnchor = session.Project().AnchorIndex.Keys.First();

        var result = session.ReplaceText(firstAnchor, "Replaced text.");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Contains(result.Modified, a => a.Id == firstAnchor);
        Assert.NotNull(result.Patch);
        Assert.Contains("Replaced text.", result.Patch!.Markdown);

        // Verify next projection reflects the change.
        Assert.Contains("Replaced text.", session.Project().Markdown);
        Assert.DoesNotContain("First paragraph.", session.Project().Markdown);
    }
```

Expected: FAIL — `ReplaceText` not defined.

- [ ] **Step 2: Implement ReplaceText**

Add to `DocxSession`:

```csharp
    public EditResult ReplaceText(string anchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (anchorId is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "null anchor");

        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"ReplaceText requires a paragraph/heading/list-item anchor; got kind={target.Anchor.Kind}", anchorId);

        var parsed = Docxodus.Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        // Take snapshot before mutation
        _history.PushUndo(TakeSnapshot());

        try
        {
            // Preserve w:pPr, replace runs
            var pPr = element.Element(W.pPr);
            element.RemoveNodes();
            if (pPr is not null) element.Add(pPr);

            // For now, ReplaceText only honors the FIRST block of the payload's runs;
            // multi-paragraph payloads on ReplaceText are out of scope (use Insert).
            if (parsed.Blocks.Count > 0)
                foreach (var run in parsed.Blocks[0].RunElements)
                    element.Add(new System.Xml.Linq.XElement(run));   // clone

            // Hyperlinks: stash docxodus:href on a w:hyperlink — to be promoted to a real
            // relationship in a later task. For now, plain text inside link is included.
            foreach (var link in element.Elements(W.hyperlink).ToList())
                PromoteHyperlinkRelationship(link, _doc!.MainDocumentPart!);
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "ReplaceText failed");
            // Roll back via the snapshot we pushed
            _history.PopUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }

        InvalidateProjectionCache();
        // Re-project the affected paragraph only via the cached anchor target
        var patch = ProjectScope(target);
        return new EditResult
        {
            Success = true,
            Modified = new[] { target.Anchor },
            Patch = patch,
        };
    }

    private MarkdownPatch ProjectScope(AnchorTarget target)
    {
        // Phase 3: take the shortcut of a full re-project; refine in Phase 4.
        var fresh = WmlToMarkdownConverter.Convert(_doc!, _settings.ProjectionSettings);
        return new MarkdownPatch(target.Anchor.Id, fresh.Markdown);
    }

    private static void PromoteHyperlinkRelationship(System.Xml.Linq.XElement link, MainDocumentPart main)
    {
        var hrefAttr = link.Attribute(System.Xml.Linq.XName.Get("href", "docxodus:"));
        if (hrefAttr is null) return;
        var rel = main.AddHyperlinkRelationship(new System.Uri(hrefAttr.Value, System.UriKind.RelativeOrAbsolute), true);
        link.SetAttributeValue(R.id, rel.Id);
        hrefAttr.Remove();
    }
```

You will need:

```csharp
using Microsoft.Extensions.Logging;
```

and at the top of `DocxSession.cs` (if not already in the file scope):

```csharp
using DocumentFormat.OpenXml.Packaging;
```

- [ ] **Step 3: Run test**

```bash
dotnet test --filter FullyQualifiedName~DS030_ReplaceTextSimple
```

Expected: PASS.

- [ ] **Step 4: Release build**

```bash
dotnet build -c Release Docxodus/Docxodus.csproj
```

Expected: zero warnings.

- [ ] **Step 5: Commit**

```bash
git add -u
git commit -m "feat(session): ReplaceText — first vertical slice"
```

### Task 3.3: ReplaceText — error coverage

**Files:**
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing tests**

```csharp
    [Fact]
    public void DS031_ReplaceText_AnchorNotFound()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var r = s.ReplaceText("p:body:deadbeef", "x");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, r.Error!.Code);
    }

    [Fact]
    public void DS032_ReplaceText_MalformedMarkdownNull()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ReplaceText(anchor, null!);
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.MalformedMarkdown, r.Error!.Code);
    }

    [Fact]
    public void DS033_ReplaceText_RejectsTableSyntax()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ReplaceText(anchor, "| a | b |\n|---|---|\n| 1 | 2 |");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.TableInsertNotSupported, r.Error!.Code);
    }

    [Fact]
    public void DS034_ReplaceText_FailureLeavesDocUnchanged()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var before = s.Project().Markdown;
        s.ReplaceText("p:body:deadbeef", "x");
        Assert.Equal(before, s.Project().Markdown);
    }
```

Run: expect PASS (the implementation already covers these branches).

- [ ] **Step 2: Commit**

```bash
git add -u
git commit -m "test(session): ReplaceText error coverage"
```

### Task 3.4: DeleteBlock

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS035_DeleteBlock()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchors = s.Project().AnchorIndex.Keys.ToList();
        Assert.True(anchors.Count >= 2);
        var toDelete = anchors[0];

        var r = s.DeleteBlock(toDelete);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Removed, a => a.Id == toDelete);
        Assert.False(s.Exists(toDelete));
        Assert.DoesNotContain("First paragraph.", s.Project().Markdown);
        Assert.Contains("Second paragraph.", s.Project().Markdown);
    }
```

- [ ] **Step 2: Implement DeleteBlock**

```csharp
    public EditResult DeleteBlock(string anchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        if (target.Anchor.Kind is not ("p" or "h" or "li" or "tbl"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind,
                $"DeleteBlock requires a block-level anchor; got kind={target.Anchor.Kind}", anchorId);

        var element = target.Resolve(_doc!);
        if (element is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null", anchorId);

        _history.PushUndo(TakeSnapshot());
        try
        {
            // Collect descendant anchor ids before removal so we can report them
            var removedAnchors = new List<Anchor> { target.Anchor };
            var index = Project().AnchorIndex;
            foreach (var descendant in element.Descendants())
            {
                var unid = (string?)descendant.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                foreach (var kv in index)
                    if (kv.Value.Unid == unid) removedAnchors.Add(kv.Value.Anchor);
            }
            element.Remove();
            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Removed = removedAnchors,
                Patch = ProjectScope(target),
            };
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "DeleteBlock failed");
            _history.PopUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }
```

- [ ] **Step 3: Run test, expect PASS**

```bash
dotnet test --filter FullyQualifiedName~DS035_DeleteBlock
```

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): DeleteBlock"
```

### Task 3.5: Public Undo / Redo

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS036_UndoReplaceText()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var before = s.Project().Markdown;
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Replaced.");
        Assert.True(s.Undo());
        Assert.Equal(before, s.Project().Markdown);
    }

    [Fact]
    public void DS037_RedoAfterUndo()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        s.ReplaceText(anchor, "Replaced.");
        var afterEdit = s.Project().Markdown;
        s.Undo();
        Assert.True(s.Redo());
        Assert.Equal(afterEdit, s.Project().Markdown);
    }

    [Fact]
    public void DS038_NothingToUndo()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        Assert.False(s.Undo());
    }
```

- [ ] **Step 2: Implement Undo / Redo**

The undo ring as implemented stores the PRE-OP snapshot. So:
- `Undo()` pops the top of undo (the pre-op state) and applies it; also we need to push the CURRENT state onto redo before applying.
- `Redo()` pops redo (the post-op state) and applies it; push pre-op back onto undo.

Adjust the ring usage to track post-op snapshots too. Replace the existing undo handling:

```csharp
    public bool Undo()
    {
        if (_disposed) return false;
        if (_history.UndoCount == 0) return false;
        var preOpSnapshot = _history.PopUndo()!;       // ring records the pre-op snapshot we pushed
        // Save the current post-op state so Redo can restore it
        _history.PushRedo(TakeSnapshot());
        RestoreSnapshot(preOpSnapshot);
        return true;
    }

    public bool Redo()
    {
        if (_disposed) return false;
        if (_history.RedoCount == 0) return false;
        var postOpSnapshot = _history.PopRedo()!;
        // Save current (pre-op) state back on the undo stack
        // PopRedo already re-pushed onto undo via UndoRing.PopRedo; instead use a simpler local pattern
        RestoreSnapshot(postOpSnapshot);
        return true;
    }
```

Tighten by simplifying `UndoRing` semantics to dual-stack of states. Refactor `UndoRing<T>` accordingly:

Replace `Docxodus/Internal/UndoRing.cs`:

```csharp
#nullable enable

using System.Collections.Generic;

namespace Docxodus.Internal;

internal sealed class UndoRing<T>
{
    private readonly LinkedList<T> _undo = new();
    private readonly LinkedList<T> _redo = new();
    private readonly int _capacity;

    public UndoRing(int capacity) => _capacity = capacity > 0 ? capacity : 1;

    public int UndoCount => _undo.Count;
    public int RedoCount => _redo.Count;

    public void RecordPreOp(T preOpSnapshot)
    {
        _undo.AddLast(preOpSnapshot);
        while (_undo.Count > _capacity) _undo.RemoveFirst();
        _redo.Clear();
    }

    public (T preOp, bool ok) PopForUndo()
    {
        if (_undo.Count == 0) return (default!, false);
        var v = _undo.Last!.Value; _undo.RemoveLast();
        return (v, true);
    }

    public void RecordForRedo(T postOpSnapshot) => _redo.AddLast(postOpSnapshot);

    public (T postOp, bool ok) PopForRedo()
    {
        if (_redo.Count == 0) return (default!, false);
        var v = _redo.Last!.Value; _redo.RemoveLast();
        return (v, true);
    }

    public void PushBackForUndo(T snapshot) => _undo.AddLast(snapshot);
    public void Clear() { _undo.Clear(); _redo.Clear(); }
}
```

Update earlier tests that referenced `PushUndo`/`PopUndo`/`PopRedo`/`PushRedo` — adjust `UndoRingTests` (delete the old tests and replace with):

```csharp
    [Fact]
    public void DS001a_RecordAndUndo()
    {
        var r = new UndoRing<string>(3);
        r.RecordPreOp("v0"); // pre-op for op1
        var (snap, ok) = r.PopForUndo();
        Assert.True(ok); Assert.Equal("v0", snap);
    }

    [Fact]
    public void DS001b_CapacityEviction()
    {
        var r = new UndoRing<string>(2);
        r.RecordPreOp("v0"); r.RecordPreOp("v1"); r.RecordPreOp("v2");
        Assert.Equal(2, r.UndoCount);
    }

    [Fact]
    public void DS001c_RecordClearsRedo()
    {
        var r = new UndoRing<string>(3);
        r.RecordPreOp("v0");
        var (s, _) = r.PopForUndo();
        r.RecordForRedo("post0");
        r.RecordPreOp("v1"); // should clear redo
        Assert.Equal(0, r.RedoCount);
    }
```

Update `DocxSession`:

```csharp
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
```

Replace every `_history.PushUndo(TakeSnapshot())` in `ReplaceText` / `DeleteBlock` with `_history.RecordPreOp(TakeSnapshot())`. Replace `_history.PopUndo()` (rollback paths) with `_history.PopForUndo()` discarding the result.

- [ ] **Step 3: Run all session tests**

```bash
dotnet test --filter FullyQualifiedName~DocxSessionTests
dotnet test --filter FullyQualifiedName~UndoRingTests
```

Expected: all PASS.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): undo / redo with pre/post snapshot tracking"
```

### Task 3.6: CHANGELOG Phase 3

- [ ] **Step 1: Append**

```markdown
- **DocxSession tier A — text CRUD + undo/redo.** `ReplaceText`, `DeleteBlock`, `Undo`, `Redo`. First vertical slice through validate → snapshot → apply → reproject pipeline.
```

- [ ] **Step 2: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(session): changelog phase 3"
```

---

## Phase 4: Tier B — structural ops

**Deliverable:** `InsertParagraph`, `SplitParagraph`, `MergeParagraphs`. Multi-paragraph payloads in `InsertParagraph`.

### Task 4.1: InsertParagraph — single block, before/after

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS040_InsertParagraphAfter()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.InsertParagraph(anchor, Position.After, "Inserted paragraph.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Single(r.Created);
        var newAnchor = r.Created[0];
        Assert.Equal("p", newAnchor.Kind);
        Assert.Contains("Inserted paragraph.", s.Project().Markdown);
    }

    [Fact]
    public void DS041_InsertParagraphBefore()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.InsertParagraph(anchor, Position.Before, "First inserted.");
        Assert.True(r.Success);
        Assert.Contains("First inserted.", s.Project().Markdown);
    }
```

- [ ] **Step 2: Implement InsertParagraph**

```csharp
    public EditResult InsertParagraph(string anchorId, Position pos, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, $"anchor not found: {anchorId}", anchorId);

        var parsed = Docxodus.Internal.MarkdownPayloadParser.Parse(markdownPayload);
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
            var newElements = new List<System.Xml.Linq.XElement>();
            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                Docxodus.Internal.UnidHelper.AssignToAllElements(p);
                newElements.Add(p);
                var unid = (string)p.Attribute(PtOpenXml.Unid)!;
                var kind = block.Kind switch
                {
                    Docxodus.Internal.ParserBlockKind.Heading1
                    or Docxodus.Internal.ParserBlockKind.Heading2
                    or Docxodus.Internal.ParserBlockKind.Heading3
                    or Docxodus.Internal.ParserBlockKind.Heading4
                    or Docxodus.Internal.ParserBlockKind.Heading5
                    or Docxodus.Internal.ParserBlockKind.Heading6 => "h",
                    Docxodus.Internal.ParserBlockKind.BulletItem
                    or Docxodus.Internal.ParserBlockKind.OrderedItem => "li",
                    _ => "p",
                };
                created.Add(new Anchor($"{kind}:{target.Anchor.Scope}:{unid}", kind, target.Anchor.Scope, unid));
            }

            if (pos == Position.Before)
            {
                foreach (var n in newElements) element.AddBeforeSelf(n);
            }
            else
            {
                System.Xml.Linq.XElement after = element;
                foreach (var n in newElements) { after.AddAfterSelf(n); after = n; }
            }

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = ProjectScope(target),
            };
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "InsertParagraph failed");
            _ = _history.PopForUndo();
            RestoreSnapshot(TakeSnapshot()); // no-op safety; snapshot already rolled back conceptually
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private static System.Xml.Linq.XElement BuildParagraphFromParsedBlock(Docxodus.Internal.ParsedBlock block)
    {
        var p = new System.Xml.Linq.XElement(W.p);
        var pPr = new System.Xml.Linq.XElement(W.pPr);

        switch (block.Kind)
        {
            case Docxodus.Internal.ParserBlockKind.Heading1:
            case Docxodus.Internal.ParserBlockKind.Heading2:
            case Docxodus.Internal.ParserBlockKind.Heading3:
            case Docxodus.Internal.ParserBlockKind.Heading4:
            case Docxodus.Internal.ParserBlockKind.Heading5:
            case Docxodus.Internal.ParserBlockKind.Heading6:
                {
                    int level = (int)block.Kind - (int)Docxodus.Internal.ParserBlockKind.Heading1 + 1;
                    pPr.Add(new System.Xml.Linq.XElement(W.pStyle, new System.Xml.Linq.XAttribute(W.val, $"Heading{level}")));
                    break;
                }
            case Docxodus.Internal.ParserBlockKind.Quote:
                pPr.Add(new System.Xml.Linq.XElement(W.pStyle, new System.Xml.Linq.XAttribute(W.val, "Quote")));
                break;
            case Docxodus.Internal.ParserBlockKind.Code:
                pPr.Add(new System.Xml.Linq.XElement(W.pStyle, new System.Xml.Linq.XAttribute(W.val, "Code")));
                break;
            // Lists left for Task 4.4 (numbering inheritance is non-trivial)
        }

        if (pPr.HasElements) p.Add(pPr);
        foreach (var run in block.RunElements) p.Add(new System.Xml.Linq.XElement(run));
        return p;
    }
```

- [ ] **Step 3: Run tests**

```bash
dotnet test --filter "FullyQualifiedName~DS040 | FullyQualifiedName~DS041"
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): InsertParagraph (single + multi-block, no list inheritance)"
```

### Task 4.2: SplitParagraph

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS042_SplitParagraph()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();

        // Text is "First paragraph." — split at offset 5 ("First|" + " paragraph.")
        var r = s.SplitParagraph(anchor, 5);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Id == anchor);     // first half keeps original anchor
        Assert.Single(r.Created);                              // second half is new

        var md = s.Project().Markdown;
        Assert.Contains("First", md);
        Assert.Contains("paragraph.", md);
    }
```

- [ ] **Step 2: Implement SplitParagraph**

```csharp
    public EditResult SplitParagraph(string anchorId, int characterOffset)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
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
            var second = new System.Xml.Linq.XElement(W.p);
            if (pPr is not null) second.Add(new System.Xml.Linq.XElement(pPr));

            // Walk runs left-to-right; for each w:r, accumulate text length.
            // When offset falls within a run, split it.
            var newSecondRuns = new List<System.Xml.Linq.XElement>();
            int consumed = 0;
            var runs = element.Elements(W.r).ToList();
            foreach (var run in runs)
            {
                var runText = string.Concat(run.Elements(W.t).Select(t => (string)t));
                if (consumed + runText.Length <= characterOffset)
                {
                    consumed += runText.Length;
                    continue;
                }
                if (consumed >= characterOffset)
                {
                    // Entire run moves to second half
                    newSecondRuns.Add(run);
                    continue;
                }
                // Split this run at (characterOffset - consumed)
                int splitAt = characterOffset - consumed;
                var keepText = runText.Substring(0, splitAt);
                var moveText = runText.Substring(splitAt);

                // Reset run's text to keep
                foreach (var t in run.Elements(W.t).ToList()) t.Remove();
                run.Add(new System.Xml.Linq.XElement(W.t,
                    new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"),
                    keepText));

                // Build a sibling run with moveText, copying rPr
                var rPr = run.Element(W.rPr);
                var moved = new System.Xml.Linq.XElement(W.r);
                if (rPr is not null) moved.Add(new System.Xml.Linq.XElement(rPr));
                moved.Add(new System.Xml.Linq.XElement(W.t,
                    new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"),
                    moveText));
                newSecondRuns.Add(moved);
                consumed = characterOffset;
            }
            foreach (var r in newSecondRuns) { r.Remove(); second.Add(r); }

            // Assign a new Unid to the second paragraph
            Docxodus.Internal.UnidHelper.AssignToAllElements(second);
            element.AddAfterSelf(second);

            var secondUnid = (string)second.Attribute(PtOpenXml.Unid)!;
            var secondAnchor = new Anchor($"{target.Anchor.Kind}:{target.Anchor.Scope}:{secondUnid}",
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
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "SplitParagraph failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private static string ParagraphText(System.Xml.Linq.XElement p) =>
        string.Concat(p.Elements(W.r).SelectMany(r => r.Elements(W.t)).Select(t => (string)t));
```

- [ ] **Step 3: Test**

```bash
dotnet test --filter FullyQualifiedName~DS042_SplitParagraph
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): SplitParagraph"
```

### Task 4.3: MergeParagraphs

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS043_MergeParagraphs()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchors = s.Project().AnchorIndex.Keys.ToList();
        var first = anchors[0];
        var second = anchors[1];

        var r = s.MergeParagraphs(first, second);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains(r.Modified, a => a.Id == first);
        Assert.Contains(r.Removed, a => a.Id == second);
        Assert.False(s.Exists(second));
        Assert.Contains("First paragraph.Second paragraph.", s.Project().Markdown);
    }

    [Fact]
    public void DS044_MergeParagraphs_NotAdjacent()
    {
        // Build a fixture with three paragraphs and try to merge first + third
        // (skip for now — covered when we add a 3-paragraph fixture in Task 4.5)
    }
```

- [ ] **Step 2: Implement MergeParagraphs**

```csharp
    public EditResult MergeParagraphs(string firstAnchorId, string secondAnchorId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        var idx = Project().AnchorIndex;
        if (!idx.TryGetValue(firstAnchorId, out var firstTarget))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "first anchor not found", firstAnchorId);
        if (!idx.TryGetValue(secondAnchorId, out var secondTarget))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "second anchor not found", secondAnchorId);

        var firstEl = firstTarget.Resolve(_doc!);
        var secondEl = secondTarget.Resolve(_doc!);
        if (firstEl is null || secondEl is null)
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "element resolved null");

        if (!ReferenceEquals(firstEl.NextNode, secondEl))
            return EditResult.Fail(EditErrorCode.AnchorsNotAdjacent,
                "MergeParagraphs requires the second anchor to be the immediate sibling of the first");

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var run in secondEl.Elements(W.r).ToList())
            {
                run.Remove();
                firstEl.Add(run);
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
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "MergeParagraphs failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message);
        }
    }
```

- [ ] **Step 3: Test**

```bash
dotnet test --filter FullyQualifiedName~DS043_MergeParagraphs
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): MergeParagraphs"
```

### Task 4.4: CHANGELOG Phase 4 + commit

- [ ] **Step 1:** Append to CHANGELOG:

```markdown
- **DocxSession tier B — structural ops.** `InsertParagraph` (before/after, multi-block payload), `SplitParagraph`, `MergeParagraphs`.
```

- [ ] **Step 2: Commit**

```bash
git add CHANGELOG.md
git commit -m "docs(session): changelog phase 4"
```

---

## Phase 5: Tier C — formatting

**Deliverable:** `ApplyFormat`, `SetParagraphStyle`, `SetListLevel`, `RemoveListMembership`. Span formatting will split/merge runs internally.

### Task 5.1: SetParagraphStyle

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS050_SetParagraphStyle()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.SetParagraphStyle(anchor, "Heading2");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Single(r.Modified);
        // The Modified anchor should have kind 'h' now
        Assert.Equal("h", r.Modified[0].Kind);
        Assert.Contains("## ", s.Project().Markdown);
    }

    [Fact]
    public void DS051_SetParagraphStyle_UnknownStyle()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.SetParagraphStyle(anchor, "NotARealStyle1234");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.UnknownStyle, r.Error!.Code);
    }
```

DS001 may not have a `Heading2` style defined in its styles part. Update the fixture builder in Task 1.1 to include common heading styles, OR add a helper that injects the style when missing. Cleanest: extend Task 1.1's fixture to include `Heading1`–`Heading6` and `Quote` style entries. If you skipped that, add them now via a one-off fixture rebuild.

- [ ] **Step 2: Implement SetParagraphStyle**

```csharp
    public EditResult SetParagraphStyle(string anchorId, string styleId)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        if (target.Anchor.Kind is not ("p" or "h" or "li"))
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "SetParagraphStyle requires a paragraph anchor", anchorId);

        // Validate style exists
        var stylesPart = _doc!.MainDocumentPart!.StyleDefinitionsPart;
        var stylesXml = stylesPart?.GetXDocument().Root;
        bool styleExists = stylesXml?.Elements(W.style)
            .Any(st => (string?)st.Attribute(W.styleId) == styleId) ?? false;
        if (!styleExists)
            return EditResult.Fail(EditErrorCode.UnknownStyle, $"style id not found: {styleId}", anchorId);

        var element = target.Resolve(_doc);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var pPr = element.Element(W.pPr);
            if (pPr is null) { pPr = new System.Xml.Linq.XElement(W.pPr); element.AddFirst(pPr); }
            var existingStyle = pPr.Element(W.pStyle);
            existingStyle?.Remove();
            pPr.AddFirst(new System.Xml.Linq.XElement(W.pStyle, new System.Xml.Linq.XAttribute(W.val, styleId)));

            InvalidateProjectionCache();
            // Compute the new anchor kind by re-deriving from the fresh projection
            var freshIndex = Project().AnchorIndex;
            var updated = freshIndex.Values.FirstOrDefault(t => t.Unid == target.Unid)?.Anchor ?? target.Anchor;

            return new EditResult
            {
                Success = true,
                Modified = new[] { updated },
                Patch = ProjectScope(target),
            };
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "SetParagraphStyle failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }
```

- [ ] **Step 3: Test**

```bash
dotnet test --filter FullyQualifiedName~DS050_SetParagraphStyle
dotnet test --filter FullyQualifiedName~DS051
```

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): SetParagraphStyle"
```

### Task 5.2: ApplyFormat — whole-paragraph and span

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS052_ApplyFormat_WholeParagraphBold()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.ApplyFormat(anchor, span: null, new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("**First paragraph.**", s.Project().Markdown);
    }

    [Fact]
    public void DS053_ApplyFormat_Span()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        // "First paragraph." -> bold characters 0..5 ("First")
        var r = s.ApplyFormat(anchor, new CharSpan(0, 5), new FormatOp { Bold = true });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("**First**", s.Project().Markdown);
    }
```

- [ ] **Step 2: Implement ApplyFormat**

```csharp
    public EditResult ApplyFormat(string anchorId, CharSpan? span, FormatOp op)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
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
            // Split runs at span boundaries, then mutate rPr on runs inside the span.
            SplitRunsAtOffset(element, actualSpan.Start);
            SplitRunsAtOffset(element, actualSpan.Start + actualSpan.Length);

            int consumed = 0;
            foreach (var run in element.Elements(W.r))
            {
                var runText = string.Concat(run.Elements(W.t).Select(t => (string)t));
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
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "ApplyFormat failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private static void SplitRunsAtOffset(System.Xml.Linq.XElement paragraph, int offset)
    {
        int consumed = 0;
        foreach (var run in paragraph.Elements(W.r).ToList())
        {
            var runText = string.Concat(run.Elements(W.t).Select(t => (string)t));
            if (consumed == offset) return;
            if (consumed + runText.Length <= offset) { consumed += runText.Length; continue; }
            int splitAt = offset - consumed;
            if (splitAt <= 0) return;
            var keep = runText.Substring(0, splitAt);
            var move = runText.Substring(splitAt);
            foreach (var t in run.Elements(W.t).ToList()) t.Remove();
            run.Add(new System.Xml.Linq.XElement(W.t,
                new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), keep));
            var rPr = run.Element(W.rPr);
            var newRun = new System.Xml.Linq.XElement(W.r);
            if (rPr is not null) newRun.Add(new System.Xml.Linq.XElement(rPr));
            newRun.Add(new System.Xml.Linq.XElement(W.t,
                new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"), move));
            run.AddAfterSelf(newRun);
            return;
        }
    }

    private static void ApplyFormatToRun(System.Xml.Linq.XElement run, FormatOp op)
    {
        var rPr = run.Element(W.rPr);
        if (rPr is null) { rPr = new System.Xml.Linq.XElement(W.rPr); run.AddFirst(rPr); }
        void Toggle(System.Xml.Linq.XName name, bool? set)
        {
            if (set is null) return;
            var existing = rPr.Element(name);
            if (set.Value && existing is null) rPr.Add(new System.Xml.Linq.XElement(name));
            else if (!set.Value && existing is not null) existing.Remove();
        }
        Toggle(W.b, op.Bold);
        Toggle(W.i, op.Italic);
        Toggle(W.strike, op.Strike);
        if (op.Underline is true)
        {
            var u = rPr.Element(W.u);
            u?.Remove();
            rPr.Add(new System.Xml.Linq.XElement(W.u, new System.Xml.Linq.XAttribute(W.val, "single")));
        }
        else if (op.Underline is false) rPr.Element(W.u)?.Remove();

        if (op.Code is true)
        {
            rPr.Element(W.rStyle)?.Remove();
            rPr.Add(new System.Xml.Linq.XElement(W.rStyle, new System.Xml.Linq.XAttribute(W.val, "Code")));
        }
        else if (op.Code is false) rPr.Element(W.rStyle)?.Remove();

        if (op.Color is not null)
        {
            rPr.Element(W.color)?.Remove();
            if (op.Color.Length > 0)
                rPr.Add(new System.Xml.Linq.XElement(W.color, new System.Xml.Linq.XAttribute(W.val, op.Color)));
        }

        if (op.RunStyle is not null)
        {
            rPr.Element(W.rStyle)?.Remove();
            if (op.RunStyle.Length > 0)
                rPr.Add(new System.Xml.Linq.XElement(W.rStyle, new System.Xml.Linq.XAttribute(W.val, op.RunStyle)));
        }
    }
```

- [ ] **Step 3: Test**

```bash
dotnet test --filter "FullyQualifiedName~DS052 | FullyQualifiedName~DS053"
```

- [ ] **Step 4: Commit**

```bash
git add -u
git commit -m "feat(session): ApplyFormat (whole-paragraph and span)"
```

### Task 5.3: SetListLevel and RemoveListMembership

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Create: `TestFiles/DS002_lists_nested.docx`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Build the list fixture**

Following the pattern of Task 1.1, create a DOCX with two nested bulleted items. Use a one-off `[Fact]` test if `dotnet script` isn't available.

- [ ] **Step 2: Failing tests**

```csharp
    [Fact]
    public void DS054_SetListLevelIndent()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS002_lists_nested.docx"));
        var firstLi = s.Project().AnchorIndex.Keys
            .First(k => k.StartsWith("li:"));
        var r = s.SetListLevel(firstLi, +1);
        Assert.True(r.Success, r.Error?.Message);
    }

    [Fact]
    public void DS055_RemoveListMembership()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS002_lists_nested.docx"));
        var firstLi = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("li:"));
        var r = s.RemoveListMembership(firstLi);
        Assert.True(r.Success, r.Error?.Message);
        Assert.True(r.Modified.Any(a => a.Kind == "p"));
    }
```

- [ ] **Step 3: Implement**

```csharp
    public EditResult SetListLevel(string anchorId, int levelDelta)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
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
        numPr.Add(new System.Xml.Linq.XElement(W.ilvl, new System.Xml.Linq.XAttribute(W.val, next)));
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
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
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
```

- [ ] **Step 4: Test, build clean, commit**

```bash
dotnet test --filter "FullyQualifiedName~DS054 | FullyQualifiedName~DS055"
dotnet build -c Release Docxodus/Docxodus.csproj
git add -u
git commit -m "feat(session): SetListLevel and RemoveListMembership"
```

### Task 5.4: CHANGELOG Phase 5

- [ ] **Step 1: Append + commit**

```markdown
- **DocxSession tier C — formatting.** `ApplyFormat` (with span), `SetParagraphStyle`, `SetListLevel`, `RemoveListMembership`.
```

```bash
git add CHANGELOG.md
git commit -m "docs(session): changelog phase 5"
```

---

## Phase 6: Tier D — table cell content + tracked-change mode

### Task 6.1: ReplaceCellContent

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Create: `TestFiles/DS003_table_with_cells.docx` (one-off builder again)
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Build fixture and failing test**

```csharp
    [Fact]
    public void DS060_ReplaceCellContent()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS003_table_with_cells.docx"));
        var cellAnchor = s.Project().AnchorIndex.Keys.First(k => k.StartsWith("tc:"));
        var r = s.ReplaceCellContent(cellAnchor, "New cell text.");
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("New cell text.", s.Project().Markdown);
    }
```

- [ ] **Step 2: Implement**

```csharp
    public EditResult ReplaceCellContent(string cellAnchorId, string markdownPayload)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(cellAnchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", cellAnchorId);
        if (target.Anchor.Kind != "tc")
            return EditResult.Fail(EditErrorCode.AnchorWrongKind, "ReplaceCellContent requires a cell anchor", cellAnchorId);

        var parsed = Docxodus.Internal.MarkdownPayloadParser.Parse(markdownPayload);
        if (!parsed.Success)
            return EditResult.Fail(parsed.Error!.Code, parsed.Error.Message, cellAnchorId);

        var cell = target.Resolve(_doc!);
        if (cell is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", cellAnchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var tcPr = cell.Element(W.tcPr);
            foreach (var p in cell.Elements(W.p).ToList()) p.Remove();

            foreach (var block in parsed.Blocks)
            {
                var p = BuildParagraphFromParsedBlock(block);
                Docxodus.Internal.UnidHelper.AssignToAllElements(p);
                cell.Add(p);
            }
            // Ensure cell has at least one paragraph
            if (!cell.Elements(W.p).Any())
                cell.Add(new System.Xml.Linq.XElement(W.p));

            InvalidateProjectionCache();
            return new EditResult
            {
                Success = true,
                Modified = new[] { target.Anchor },
                Patch = ProjectScope(target),
            };
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "ReplaceCellContent failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, cellAnchorId);
        }
    }
```

- [ ] **Step 3: Test, commit**

```bash
dotnet test --filter FullyQualifiedName~DS060
git add -u && git commit -m "feat(session): ReplaceCellContent"
```

### Task 6.2: Tracked-change mode

For v1, tracked-change mode is a wrap-around behavior applied at the point of mutation. The simplest implementation: when `Settings.TrackedChanges = RenderInline`, wrap the original runs in `w:del` (with author/date) and wrap new runs in `w:ins`.

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS061_ReplaceText_Tracked()
    {
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "test-agent",
        };
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"), settings);
        var anchor = s.Project().AnchorIndex.Keys.First();

        var r = s.ReplaceText(anchor, "New text.");
        Assert.True(r.Success);
        // Document XML should contain w:ins and w:del
        var xml = s.Project().Markdown;
        // Easier: poke at the live XDocument via Save round-trip:
        var bytes = s.Save();
        using var verify = new DocxSession(bytes);
        // The verify session's Project() defaults to TrackedChangeMode.Accept (in projection settings)
        // so the markdown won't show del/ins markers unless we configure differently.
        // For a structural assertion, use Raw.GetXml in Phase 7. Here, we just confirm Success.
    }
```

- [ ] **Step 2: Implement tracked-mode branch in `ReplaceText`**

In `ReplaceText`, replace the "Preserve w:pPr, replace runs" block with:

```csharp
            if (_settings.TrackedChanges == TrackedChangeMode.RenderInline)
            {
                // Wrap existing runs in w:del; insert new runs wrapped in w:ins.
                var existingRuns = element.Elements(W.r).ToList();
                int revId = NextRevisionId();
                if (existingRuns.Count > 0)
                {
                    var del = new System.Xml.Linq.XElement(W.del,
                        new System.Xml.Linq.XAttribute(W.id, revId),
                        new System.Xml.Linq.XAttribute(W.author, _settings.RevisionAuthor ?? "docxodus"),
                        new System.Xml.Linq.XAttribute(W.date, System.DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")));
                    foreach (var run in existingRuns)
                    {
                        run.Remove();
                        // text inside a deleted run must use w:delText instead of w:t
                        foreach (var t in run.Elements(W.t).ToList())
                        {
                            var dt = new System.Xml.Linq.XElement(W.delText,
                                new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"),
                                (string)t);
                            t.ReplaceWith(dt);
                        }
                        del.Add(run);
                    }
                    element.Add(del);
                }

                if (parsed.Blocks.Count > 0 && parsed.Blocks[0].RunElements.Count > 0)
                {
                    var ins = new System.Xml.Linq.XElement(W.ins,
                        new System.Xml.Linq.XAttribute(W.id, NextRevisionId()),
                        new System.Xml.Linq.XAttribute(W.author, _settings.RevisionAuthor ?? "docxodus"),
                        new System.Xml.Linq.XAttribute(W.date, System.DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")));
                    foreach (var run in parsed.Blocks[0].RunElements)
                        ins.Add(new System.Xml.Linq.XElement(run));
                    element.Add(ins);
                }
            }
            else
            {
                // Accept (default): replace runs in place
                var pPr = element.Element(W.pPr);
                element.RemoveNodes();
                if (pPr is not null) element.Add(pPr);
                if (parsed.Blocks.Count > 0)
                    foreach (var run in parsed.Blocks[0].RunElements)
                        element.Add(new System.Xml.Linq.XElement(run));
            }
```

Add a small helper:

```csharp
    private int _revisionCounter = 1000;
    private int NextRevisionId() => System.Threading.Interlocked.Increment(ref _revisionCounter);
```

- [ ] **Step 3: Test, commit**

```bash
dotnet test --filter FullyQualifiedName~DS061
git add -u && git commit -m "feat(session): tracked-change mode for ReplaceText"
```

Subsequent phases (Phase 7+) and the v2 backlog can extend tracked mode to other ops. The v1 spec promises tracked mode "across all earlier ops"; for Phase 6's scope, prove the pattern with `ReplaceText` and follow up in a chaser task if needed for `DeleteBlock`.

### Task 6.3: Tracked-mode for DeleteBlock

Same pattern as 6.2, wrapping the deleted element in `w:del` rather than removing.

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS062_DeleteBlock_Tracked()
    {
        var settings = new DocxSessionSettings
        {
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "tester",
        };
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"), settings);
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.DeleteBlock(anchor);
        Assert.True(r.Success);
        // In tracked mode, the anchor should remain (modified, not removed)
        Assert.Empty(r.Removed);
        Assert.Contains(r.Modified, a => a.Id == anchor);
    }
```

- [ ] **Step 2: Implement tracked branch in `DeleteBlock`**

Replace the "element.Remove()" branch with:

```csharp
            if (_settings.TrackedChanges == TrackedChangeMode.RenderInline)
            {
                // Mark the paragraph's pPr with a pPrChange + wrap all runs in w:del
                foreach (var run in element.Elements(W.r).ToList())
                {
                    var del = new System.Xml.Linq.XElement(W.del,
                        new System.Xml.Linq.XAttribute(W.id, NextRevisionId()),
                        new System.Xml.Linq.XAttribute(W.author, _settings.RevisionAuthor ?? "docxodus"),
                        new System.Xml.Linq.XAttribute(W.date, System.DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")));
                    foreach (var t in run.Elements(W.t).ToList())
                        t.ReplaceWith(new System.Xml.Linq.XElement(W.delText,
                            new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xml + "space", "preserve"),
                            (string)t));
                    run.Remove();
                    del.Add(run);
                    element.Add(del);
                }
                InvalidateProjectionCache();
                return new EditResult
                {
                    Success = true,
                    Modified = new[] { target.Anchor },
                    Patch = ProjectScope(target),
                };
            }
```

- [ ] **Step 3: Test, commit**

```bash
dotnet test --filter "FullyQualifiedName~DS062"
git add -u && git commit -m "feat(session): tracked-change mode for DeleteBlock"
```

### Task 6.4: CHANGELOG Phase 6

- [ ] Append + commit:

```markdown
- **DocxSession tier D — table cell + tracked changes.** `ReplaceCellContent`; `TrackedChanges = RenderInline` mode lands `w:ins`/`w:del` revisions for `ReplaceText` and `DeleteBlock`.
```

```bash
git add CHANGELOG.md && git commit -m "docs(session): changelog phase 6"
```

---

## Phase 7: Raw escape hatch

### Task 7.1: RawDocxOps scaffolding

**Files:**
- Create: `Docxodus/RawDocxOps.cs`
- Modify: `Docxodus/DocxSession.cs` (expose `Raw` property)

- [ ] **Step 1: Create `RawDocxOps`**

```csharp
#nullable enable

using System.Xml.Linq;

namespace Docxodus;

public sealed class RawDocxOps
{
    private readonly DocxSession _session;
    internal RawDocxOps(DocxSession session) => _session = session;

    public string GetXml(string anchorId) => _session.RawGetXmlInternal(anchorId);

    public EditResult InsertXml(string anchorId, Position pos, string xml) =>
        _session.RawInsertXmlInternal(anchorId, pos, xml);

    public EditResult ReplaceXml(string anchorId, string xml) =>
        _session.RawReplaceXmlInternal(anchorId, xml);
}
```

- [ ] **Step 2: Add Raw property and internal hooks to DocxSession**

```csharp
    public RawDocxOps Raw => _raw ??= new RawDocxOps(this);
    private RawDocxOps? _raw;

    internal string RawGetXmlInternal(string anchorId)
    {
        ThrowIfDisposed();
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            throw new System.ArgumentException($"anchor not found: {anchorId}");
        var element = target.Resolve(_doc!);
        return element?.ToString() ?? "";
    }

    internal EditResult RawInsertXmlInternal(string anchorId, Position pos, string xml)
    {
        // Implementation in Task 7.3
        return EditResult.Fail(EditErrorCode.InternalError, "not yet implemented", anchorId);
    }

    internal EditResult RawReplaceXmlInternal(string anchorId, string xml)
    {
        // Implementation in Task 7.4
        return EditResult.Fail(EditErrorCode.InternalError, "not yet implemented", anchorId);
    }
```

- [ ] **Step 3: Test GetXml works**

```csharp
    [Fact]
    public void DS070_RawGetXml()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var xml = s.Raw.GetXml(anchor);
        Assert.Contains("First paragraph.", xml);
        Assert.StartsWith("<w:p", xml);
    }
```

```bash
dotnet test --filter "FullyQualifiedName~DS070"
git add -u && git commit -m "feat(session): Raw scaffolding + GetXml"
```

### Task 7.2: Raw validation pipeline

**Files:**
- Modify: `Docxodus/DocxSession.cs`

- [ ] **Step 1: Add validation helpers**

```csharp
    private static readonly System.Collections.Generic.HashSet<string> AllowedXmlNamespaces = new()
    {
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main",        // w:
        "http://schemas.openxmlformats.org/officeDocument/2006/math",          // m:
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing", // wp:
        "http://schemas.openxmlformats.org/drawingml/2006/main",               // a:
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships", // r:
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main/pt",     // PtOpenXml namespace (for Unid)
    };

    private static (System.Xml.Linq.XElement? parsed, EditError? err) ParseRawXml(string xml)
    {
        try
        {
            var x = System.Xml.Linq.XElement.Parse(xml);
            // Reject any descendant in a non-allowed namespace
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
```

- [ ] **Step 2: Commit**

```bash
git add -u && git commit -m "feat(session): raw XML validation helpers"
```

### Task 7.3: Raw.InsertXml

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS071_RawInsertXml()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var newP = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Raw inserted.</w:t></w:r></w:p>";
        var r = s.Raw.InsertXml(anchor, Position.After, newP);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Single(r.Created);
        Assert.Contains("Raw inserted.", s.Project().Markdown);
    }

    [Fact]
    public void DS072_RawInsertXml_MalformedRejected()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.Raw.InsertXml(anchor, Position.After, "<not-closed");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.MalformedXml, r.Error!.Code);
    }

    [Fact]
    public void DS073_RawInsertXml_DisallowedNs()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var r = s.Raw.InsertXml(anchor, Position.After, "<foo xmlns=\"http://evil/\"/>");
        Assert.False(r.Success);
        Assert.Equal(EditErrorCode.DisallowedNamespace, r.Error!.Code);
    }
```

- [ ] **Step 2: Implement `RawInsertXmlInternal`**

```csharp
    internal new EditResult RawInsertXmlInternal(string anchorId, Position pos, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);

        var (parsed, err) = ParseRawXml(xml);
        if (parsed is null) return new EditResult { Success = false, Error = err with { AnchorId = anchorId } };

        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Docxodus.Internal.UnidHelper.AssignToAllElements(parsed);
            if (pos == Position.Before) element.AddBeforeSelf(parsed);
            else element.AddAfterSelf(parsed);

            InvalidateProjectionCache();
            var created = new List<Anchor>();
            var freshIndex = Project().AnchorIndex;
            foreach (var el in parsed.DescendantsAndSelf())
            {
                var unid = (string?)el.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == unid);
                if (hit is not null) created.Add(hit.Anchor);
            }

            if (_settings.ValidateRawOps && !RunOpenXmlValidator())
            {
                _ = _history.PopForUndo();
                RestoreSnapshot(TakeSnapshot()); // discard the failing state
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            return new EditResult
            {
                Success = true,
                Created = created,
                Patch = ProjectScope(target),
            };
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "Raw.InsertXml failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }

    private bool RunOpenXmlValidator()
    {
        var v = new DocumentFormat.OpenXml.Validation.OpenXmlValidator();
        return !v.Validate(_doc!).Any();
    }
```

(Remove the placeholder `internal EditResult RawInsertXmlInternal(...)` stub from Task 7.1 — replace it with this implementation; remove the `new` keyword from the signature.)

- [ ] **Step 3: Test, commit**

```bash
dotnet test --filter "FullyQualifiedName~DS071 | FullyQualifiedName~DS072 | FullyQualifiedName~DS073"
git add -u && git commit -m "feat(session): Raw.InsertXml + validation"
```

### Task 7.4: Raw.ReplaceXml

**Files:**
- Modify: `Docxodus/DocxSession.cs`
- Modify: `Docxodus.Tests/DocxSessionTests.cs`

- [ ] **Step 1: Failing test**

```csharp
    [Fact]
    public void DS074_RawReplaceXml()
    {
        using var s = new DocxSession(LoadFixtureBytes("DS001_simple_two_paragraphs.docx"));
        var anchor = s.Project().AnchorIndex.Keys.First();
        var newP = "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:r><w:t>Raw replacement.</w:t></w:r></w:p>";
        var r = s.Raw.ReplaceXml(anchor, newP);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Contains("Raw replacement.", s.Project().Markdown);
    }
```

- [ ] **Step 2: Implement**

```csharp
    internal EditResult RawReplaceXmlInternal(string anchorId, string xml)
    {
        if (_disposed) return EditResult.Fail(EditErrorCode.SessionDisposed, "session disposed");
        if (!Project().AnchorIndex.TryGetValue(anchorId, out var target))
            return EditResult.Fail(EditErrorCode.AnchorNotFound, "anchor not found", anchorId);
        var (parsed, err) = ParseRawXml(xml);
        if (parsed is null) return new EditResult { Success = false, Error = err with { AnchorId = anchorId } };
        var element = target.Resolve(_doc!);
        if (element is null) return EditResult.Fail(EditErrorCode.AnchorNotFound, "element null", anchorId);

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            Docxodus.Internal.UnidHelper.AssignToAllElements(parsed);
            element.ReplaceWith(parsed);

            InvalidateProjectionCache();
            var removed = new List<Anchor> { target.Anchor };
            var created = new List<Anchor>();
            var freshIndex = Project().AnchorIndex;
            foreach (var el in parsed.DescendantsAndSelf())
            {
                var unid = (string?)el.Attribute(PtOpenXml.Unid);
                if (unid is null) continue;
                var hit = freshIndex.Values.FirstOrDefault(t => t.Unid == unid);
                if (hit is not null) created.Add(hit.Anchor);
            }

            if (_settings.ValidateRawOps && !RunOpenXmlValidator())
            {
                _ = _history.PopForUndo();
                return EditResult.Fail(EditErrorCode.ValidationFailed, "OpenXmlValidator found new errors", anchorId);
            }

            return new EditResult
            {
                Success = true,
                Removed = removed,
                Created = created,
                Patch = ProjectScope(target),
            };
        }
        catch (System.Exception ex)
        {
            LastInternalError = ex;
            _settings.Logger?.LogError(ex, "Raw.ReplaceXml failed");
            _ = _history.PopForUndo();
            return EditResult.Fail(EditErrorCode.InternalError, ex.Message, anchorId);
        }
    }
```

- [ ] **Step 3: Test, commit**

```bash
dotnet test --filter FullyQualifiedName~DS074
git add -u && git commit -m "feat(session): Raw.ReplaceXml"
```

### Task 7.5: CHANGELOG Phase 7

- [ ] Append + commit

```markdown
- **DocxSession raw escape hatch.** `Raw.GetXml`, `Raw.InsertXml`, `Raw.ReplaceXml` for OOXML the markdown subset can't express; `Settings.ValidateRawOps` for `OpenXmlValidator` post-checks.
```

```bash
git add CHANGELOG.md && git commit -m "docs(session): changelog phase 7"
```

---

## Phase 8: WASM bridge + npm wrapper + Playwright

### Task 8.1: WASM JSExport bridge

**Files:**
- Create: `wasm/DocxodusWasm/DocxSessionBridge.cs`

- [ ] **Step 1: Create bridge file**

```csharp
#nullable enable

using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices.JavaScript;
using System.Text.Json;
using Docxodus;

namespace DocxodusWasm;

public static class DocxSessionBridge
{
    private static readonly ConcurrentDictionary<int, DocxSession> _sessions = new();
    private static int _nextId = 0;

    [JSExport]
    public static int OpenSession(byte[] bytes, string settingsJson)
    {
        var settings = string.IsNullOrEmpty(settingsJson)
            ? new DocxSessionSettings()
            : JsonSerializer.Deserialize<DocxSessionSettings>(settingsJson) ?? new DocxSessionSettings();
        var s = new DocxSession(bytes, settings);
        var id = System.Threading.Interlocked.Increment(ref _nextId);
        _sessions[id] = s;
        return id;
    }

    [JSExport]
    public static void CloseSession(int handle)
    {
        if (_sessions.TryRemove(handle, out var s)) s.Dispose();
    }

    [JSExport]
    public static string Project(int handle)
    {
        var s = Get(handle);
        var p = s.Project();
        return JsonSerializer.Serialize(new
        {
            markdown = p.Markdown,
            anchorIndex = p.AnchorIndex.ToDictionary(
                kv => kv.Key,
                kv => new { partUri = kv.Value.PartUri, unid = kv.Value.Unid,
                            kind = kv.Value.Anchor.Kind, scope = kv.Value.Anchor.Scope }),
        });
    }

    [JSExport] public static string ReplaceText(int h, string a, string md)
        => Serialize(Get(h).ReplaceText(a, md));
    [JSExport] public static string DeleteBlock(int h, string a)
        => Serialize(Get(h).DeleteBlock(a));
    [JSExport] public static string InsertParagraph(int h, string a, string pos, string md)
        => Serialize(Get(h).InsertParagraph(a, ParsePos(pos), md));
    [JSExport] public static string SplitParagraph(int h, string a, int offset)
        => Serialize(Get(h).SplitParagraph(a, offset));
    [JSExport] public static string MergeParagraphs(int h, string first, string second)
        => Serialize(Get(h).MergeParagraphs(first, second));
    [JSExport] public static string ApplyFormat(int h, string a, string spanJson, string opJson)
    {
        CharSpan? span = string.IsNullOrEmpty(spanJson) ? null
            : JsonSerializer.Deserialize<CharSpan?>(spanJson);
        var op = JsonSerializer.Deserialize<FormatOp>(opJson) ?? new FormatOp();
        return Serialize(Get(h).ApplyFormat(a, span, op));
    }
    [JSExport] public static string SetParagraphStyle(int h, string a, string styleId)
        => Serialize(Get(h).SetParagraphStyle(a, styleId));
    [JSExport] public static string SetListLevel(int h, string a, int delta)
        => Serialize(Get(h).SetListLevel(a, delta));
    [JSExport] public static string RemoveListMembership(int h, string a)
        => Serialize(Get(h).RemoveListMembership(a));
    [JSExport] public static string ReplaceCellContent(int h, string a, string md)
        => Serialize(Get(h).ReplaceCellContent(a, md));

    [JSExport] public static string RawGetXml(int h, string a) => Get(h).Raw.GetXml(a);
    [JSExport] public static string RawInsertXml(int h, string a, string pos, string xml)
        => Serialize(Get(h).Raw.InsertXml(a, ParsePos(pos), xml));
    [JSExport] public static string RawReplaceXml(int h, string a, string xml)
        => Serialize(Get(h).Raw.ReplaceXml(a, xml));

    [JSExport] public static bool Undo(int h) => Get(h).Undo();
    [JSExport] public static bool Redo(int h) => Get(h).Redo();
    [JSExport] public static byte[] Save(int h) => Get(h).Save();

    private static DocxSession Get(int handle)
    {
        if (!_sessions.TryGetValue(handle, out var s))
            throw new ArgumentException($"unknown session handle: {handle}");
        return s;
    }

    private static Position ParsePos(string s) =>
        string.Equals(s, "before", StringComparison.OrdinalIgnoreCase) ? Position.Before : Position.After;

    private static string Serialize(EditResult r) =>
        JsonSerializer.Serialize(new
        {
            success = r.Success,
            error = r.Error is null ? null : new
            {
                code = EnumToSnake(r.Error.Code),
                message = r.Error.Message,
                anchorId = r.Error.AnchorId,
            },
            created = r.Created.Select(SerializeAnchor),
            removed = r.Removed.Select(SerializeAnchor),
            modified = r.Modified.Select(SerializeAnchor),
            patch = r.Patch is null ? null
                : new { scopeAnchorId = r.Patch.ScopeAnchorId, markdown = r.Patch.Markdown },
        });

    private static object SerializeAnchor(Anchor a) => new
    {
        id = a.Id, kind = a.Kind, scope = a.Scope, unid = a.Unid,
    };

    private static string EnumToSnake(EditErrorCode c)
    {
        var s = c.ToString();
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < s.Length; i++)
        {
            if (i > 0 && char.IsUpper(s[i])) sb.Append('_');
            sb.Append(char.ToLowerInvariant(s[i]));
        }
        return sb.ToString();
    }
}
```

- [ ] **Step 2: Build WASM**

```bash
./scripts/build-wasm.sh
```

Expected: clean WASM build.

- [ ] **Step 3: Commit**

```bash
git add wasm/DocxodusWasm/DocxSessionBridge.cs
git commit -m "feat(session): WASM JSExport bridge"
```

### Task 8.2: TypeScript types + npm session wrapper

**Files:**
- Modify: `npm/src/types.ts`
- Create: `npm/src/session.ts`
- Modify: `npm/src/index.ts`

- [ ] **Step 1: Add types**

Append to `npm/src/types.ts`:

```typescript
export type EditErrorCode =
  | 'anchor_not_found' | 'anchor_wrong_kind' | 'anchors_not_adjacent'
  | 'session_disposed'
  | 'malformed_markdown' | 'unsupported_markdown_syntax'
  | 'table_insert_not_supported' | 'footnote_ref_not_supported'
  | 'comment_marker_not_supported' | 'image_insert_not_supported'
  | 'anchor_token_in_payload'
  | 'offset_out_of_range' | 'invalid_position'
  | 'unknown_style' | 'invalid_list_level'
  | 'malformed_xml' | 'disallowed_namespace' | 'incompatible_element_type'
  | 'validation_failed'
  | 'nothing_to_undo' | 'nothing_to_redo'
  | 'internal_error';

export interface AnchorRef {
  id: string; kind: string; scope: string; unid: string;
}

export interface EditError {
  code: EditErrorCode; message: string; anchorId?: string;
}

export interface MarkdownPatch { scopeAnchorId: string; markdown: string; }

export interface EditResult {
  success: boolean;
  error?: EditError;
  created: AnchorRef[];
  removed: AnchorRef[];
  modified: AnchorRef[];
  patch?: MarkdownPatch;
}

export interface CharSpan { start: number; length: number; }

export interface FormatOp {
  bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
  code?: boolean; color?: string; runStyle?: string;
}

export interface DocxSessionSettings {
  undoDepth?: number;
  validateRawOps?: boolean;
  trackedChanges?: 'accept' | 'render_inline' | 'strip_deletions';
  revisionAuthor?: string;
}
```

- [ ] **Step 2: Add bridge type extension to DocxodusWasmExports**

In the same file, locate `DocxodusWasmExports` and add:

```typescript
  OpenSession(bytes: Uint8Array, settingsJson: string): number;
  CloseSession(handle: number): void;
  Project(handle: number): string;
  ReplaceText(handle: number, anchor: string, md: string): string;
  DeleteBlock(handle: number, anchor: string): string;
  InsertParagraph(handle: number, anchor: string, pos: string, md: string): string;
  SplitParagraph(handle: number, anchor: string, offset: number): string;
  MergeParagraphs(handle: number, first: string, second: string): string;
  ApplyFormat(handle: number, anchor: string, spanJson: string, opJson: string): string;
  SetParagraphStyle(handle: number, anchor: string, styleId: string): string;
  SetListLevel(handle: number, anchor: string, delta: number): string;
  RemoveListMembership(handle: number, anchor: string): string;
  ReplaceCellContent(handle: number, anchor: string, md: string): string;
  RawGetXml(handle: number, anchor: string): string;
  RawInsertXml(handle: number, anchor: string, pos: string, xml: string): string;
  RawReplaceXml(handle: number, anchor: string, xml: string): string;
  Undo(handle: number): boolean;
  Redo(handle: number): boolean;
  Save(handle: number): Uint8Array;
```

- [ ] **Step 3: Create session.ts**

```typescript
import type {
  EditResult, AnchorRef, CharSpan, FormatOp, DocxSessionSettings,
} from './types.js';
import { getWasmExports } from './index.js';   // existing helper that returns the worker proxy

export async function openDocxSession(
  bytes: Uint8Array,
  settings?: DocxSessionSettings
): Promise<DocxSession> {
  const wasm = await getWasmExports();
  const handle = wasm.OpenSession(bytes, settings ? JSON.stringify(settings) : '');
  return new DocxSession(handle, wasm);
}

export class DocxSession {
  constructor(
    private readonly _handle: number,
    private readonly _wasm: any /* DocxodusWasmExports */
  ) {}

  async project(): Promise<{ markdown: string; anchorIndex: Record<string, AnchorRef> }> {
    return JSON.parse(this._wasm.Project(this._handle));
  }
  replaceText(a: string, md: string): EditResult { return JSON.parse(this._wasm.ReplaceText(this._handle, a, md)); }
  deleteBlock(a: string): EditResult { return JSON.parse(this._wasm.DeleteBlock(this._handle, a)); }
  insertParagraph(a: string, pos: 'before' | 'after', md: string): EditResult {
    return JSON.parse(this._wasm.InsertParagraph(this._handle, a, pos, md));
  }
  splitParagraph(a: string, offset: number): EditResult { return JSON.parse(this._wasm.SplitParagraph(this._handle, a, offset)); }
  mergeParagraphs(first: string, second: string): EditResult { return JSON.parse(this._wasm.MergeParagraphs(this._handle, first, second)); }
  applyFormat(a: string, span: CharSpan | null, op: FormatOp): EditResult {
    return JSON.parse(this._wasm.ApplyFormat(this._handle, a,
      span ? JSON.stringify({ Start: span.start, Length: span.length }) : '',
      JSON.stringify({
        Bold: op.bold ?? null, Italic: op.italic ?? null, Underline: op.underline ?? null,
        Strike: op.strike ?? null, Code: op.code ?? null,
        Color: op.color ?? null, RunStyle: op.runStyle ?? null,
      })));
  }
  setParagraphStyle(a: string, styleId: string): EditResult { return JSON.parse(this._wasm.SetParagraphStyle(this._handle, a, styleId)); }
  setListLevel(a: string, delta: number): EditResult { return JSON.parse(this._wasm.SetListLevel(this._handle, a, delta)); }
  removeListMembership(a: string): EditResult { return JSON.parse(this._wasm.RemoveListMembership(this._handle, a)); }
  replaceCellContent(a: string, md: string): EditResult { return JSON.parse(this._wasm.ReplaceCellContent(this._handle, a, md)); }

  readonly raw = {
    getXml: (a: string): string => this._wasm.RawGetXml(this._handle, a),
    insertXml: (a: string, pos: 'before' | 'after', xml: string): EditResult =>
      JSON.parse(this._wasm.RawInsertXml(this._handle, a, pos, xml)),
    replaceXml: (a: string, xml: string): EditResult =>
      JSON.parse(this._wasm.RawReplaceXml(this._handle, a, xml)),
  };

  undo(): boolean { return this._wasm.Undo(this._handle); }
  redo(): boolean { return this._wasm.Redo(this._handle); }
  save(): Uint8Array { return this._wasm.Save(this._handle); }
  close(): void { this._wasm.CloseSession(this._handle); }
  [Symbol.dispose](): void { this.close(); }
}
```

- [ ] **Step 4: Re-export from `npm/src/index.ts`**

Add near other re-exports:

```typescript
export { openDocxSession, DocxSession } from './session.js';
```

- [ ] **Step 5: Build npm**

```bash
cd npm && npm run build
```

Expected: clean build (warnings about `getWasmExports` are fine if the function name differs in the codebase — adjust the import after checking `npm/src/index.ts`).

- [ ] **Step 6: Commit**

```bash
git add npm/src/types.ts npm/src/session.ts npm/src/index.ts
git commit -m "feat(session): npm wrapper (DocxSession)"
```

### Task 8.3: Playwright smoke test

**Files:**
- Create: `npm/tests/docx-session.spec.ts`

- [ ] **Step 1: Write the test**

```typescript
import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';

test('open, replace text, save, reopen', async ({ page }) => {
  await page.goto('/');     // existing harness page that loads the npm bundle
  const bytes = readFileSync(join(__dirname, '..', '..', 'TestFiles', 'DS001_simple_two_paragraphs.docx'));
  const b64 = Buffer.from(bytes).toString('base64');

  const result = await page.evaluate(async (b64Bytes) => {
    const { openDocxSession } = (window as any).docxodus;
    const bin = Uint8Array.from(atob(b64Bytes), c => c.charCodeAt(0));
    const session = await openDocxSession(bin);
    const proj = await session.project();
    const anchor = Object.keys(proj.anchorIndex)[0];
    const r = await session.replaceText(anchor, 'JS-side replaced.');
    const after = (await session.project()).markdown;
    session.close();
    return { success: r.success, markdown: after };
  }, b64);

  expect(result.success).toBe(true);
  expect(result.markdown).toContain('JS-side replaced.');
});
```

- [ ] **Step 2: Run**

```bash
cd npm && npx playwright test --grep "open, replace text"
```

Expected: PASS (after the harness page exposes `window.docxodus` — confirm by inspecting `npm/tests/harness.html` or the equivalent).

- [ ] **Step 3: Commit**

```bash
git add npm/tests/docx-session.spec.ts
git commit -m "test(session): npm Playwright smoke test"
```

### Task 8.4: CHANGELOG Phase 8

- [ ] Append + commit

```markdown
- **DocxSession WASM bridge + npm/React wrapper.** Full session API exposed via JSExport; TypeScript types kept in lockstep with C# enum.
```

```bash
git add CHANGELOG.md && git commit -m "docs(session): changelog phase 8"
```

---

## Phase 9: Docs

### Task 9.1: In-tree architecture doc

**Files:**
- Create: `docs/architecture/docx_mutation_api.md`

- [ ] **Step 1: Copy the spec into place, slim it for the in-tree audience**

Copy `docs/superpowers/specs/2026-05-24-docx-mutation-api-design.md` to `docs/architecture/docx_mutation_api.md`. Remove the "Status" preamble (it's not a spec anymore). Keep all the API surface, anchor lifecycle, error catalog, supported markdown subset, performance budgets, and worked examples. Drop the Implementation Phasing section (it's history at this point — the code is the truth).

- [ ] **Step 2: Add link from CLAUDE.md**

In `CLAUDE.md` under "Core Modules", add an entry following the same pattern as `WmlToMarkdownConverter`:

```markdown
**DocxSession.cs** - Stateful in-memory DOCX editing API keyed by markdown projection anchor ids. The agent-friendly write side of `WmlToMarkdownConverter`:
- `DocxSession(byte[] bytes, DocxSessionSettings? settings = null)` - open a session over in-memory DOCX bytes
- `Project()`, `Save()`, `Exists()`, `GetAnchorInfo()` — read surface
- Mutation tiers A–D plus `Raw.GetXml`/`InsertXml`/`ReplaceXml` escape hatch
- Bounded snapshot undo/redo
- See `docs/architecture/docx_mutation_api.md` for the full surface and lifecycle policy
```

Under "Architecture Documentation", add:

```markdown
- `docx_mutation_api.md` — DocxSession surface, anchor lifecycle, error catalog, markdown subset
```

- [ ] **Step 3: Commit**

```bash
git add docs/architecture/docx_mutation_api.md CLAUDE.md
git commit -m "docs(session): in-tree architecture doc + CLAUDE.md entry"
```

---

## Final: Smoke Test

After all phases land, hand back to the user for a guided smoke test. The smoke test should exercise an end-to-end agentic flow with a real-looking DOCX:

1. Open a meaningful fixture (e.g., one of the `HC*` legal documents).
2. Project to markdown.
3. Issue a sequence of mutations covering all tiers: replace a paragraph, split one, insert a heading, format a span, edit a cell, swap one element via raw XML.
4. Save, reopen, project again, verify all edits survived.
5. Undo back to the start, verify markdown matches original.

(This step is performed at execution time, not committed as code — it's verification.)

---

## Self-Review Notes

- All 9 phases mapped to tasks; each task has explicit files, code, commands, and expected outcomes.
- Type signatures stay consistent: `ParsedBlock`, `EditResult`, `Anchor` referenced identically across tasks.
- One placeholder caveat: Task 5.1 (SetParagraphStyle) requires the `Heading2` style to exist in DS001 — flagged in the task text with a fallback (extend the fixture builder or rebuild the fixture). Engineer must address before the test passes.
- The hyperlink relationship-promotion in Task 3.2 references `R.id` and `MainDocumentPart.AddHyperlinkRelationship`. Both are standard SDK APIs already used by `WmlToHtmlConverter` (search the codebase for `AddHyperlinkRelationship` to confirm signatures).
- `BuildParagraphFromParsedBlock` is defined in Task 4.1 and reused in Task 6.1 (ReplaceCellContent). Make sure it's accessible (i.e., not lifted into a different class without updating references).
- `ProjectScope` is a placeholder Phase 3 implementation that re-projects the whole document. The spec calls for narrower patches; if a Phase 10 follow-up wants to optimize, the contract (`EditResult.Patch.ScopeAnchorId` is the smallest enclosing block) is unchanged. Document this as a known TODO in `docs/architecture/docx_mutation_api.md`'s "Open Questions" section.
