#nullable enable

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for <see cref="WmlToMarkdownConverter"/>. Test IDs follow the <c>MD###</c> prefix
/// convention documented in <c>docs/architecture/markdown_projection.md</c>. Phase numbering
/// in the doc maps to test ID ranges: phase 1 (anchor index) = MD001-MD009, phase 2
/// (paragraphs/headings) = MD010-MD019, phase 3 (inline runs) = MD020-MD029, phase 4 (lists)
/// = MD030-MD039, phase 5 (tables) = MD040-MD049, phase 6 (multipart) = MD050-MD059, phase 7
/// (tracked changes) = MD060-MD069.
/// </summary>
public class WmlToMarkdownConverterTests
{
    private static readonly DirectoryInfo TestFilesDir = new("../../../../TestFiles/");

    private static WmlDocument LoadFixture(string fixtureName) =>
        new WmlDocument(Path.Combine(TestFilesDir.FullName, fixtureName));

    // ----- Phase 1: anchor index -----

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    [InlineData("HC004-ResumeTemplate.docx")]
    public void MD001_AnchorIndexIsExhaustive(string fixtureName)
    {
        var doc = LoadFixture(fixtureName);
        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var body = wdoc.MainDocumentPart!.Document.Body!;
        var expectedBlockCount = body.Descendants()
            .Count(d => d.LocalName == "p" || d.LocalName == "tbl");

        Assert.NotEmpty(projection.AnchorIndex);

        var bodyAnchors = projection.AnchorIndex.Values.Count(t => t.Anchor.Scope == "body");
        Assert.True(bodyAnchors >= expectedBlockCount,
            $"Expected at least {expectedBlockCount} body anchors, got {bodyAnchors}");
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD002_AnchorsAreStable(string fixtureName)
    {
        // Projecting the SAME document twice MUST produce the same anchor ids — the first
        // projection persists Unids into the byte array, so the second projection reuses them.
        // (Two cold loads of the same bytes legitimately get different Unids — the fixture has
        // none stored — and that is not what "stable" means in the spec.)
        var doc = LoadFixture(fixtureName);
        var p1 = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());
        var p2 = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        Assert.Equal(
            p1.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray(),
            p2.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray());
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD003_AnchorsResolve(string fixtureName)
    {
        var doc = LoadFixture(fixtureName);
        var projection = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        using var stream = new MemoryStream(doc.DocumentByteArray);
        using var wdoc = WordprocessingDocument.Open(stream, false);

        Assert.NotEmpty(projection.AnchorIndex);
        foreach (var kvp in projection.AnchorIndex)
        {
            var target = kvp.Value;
            var element = target.Resolve(wdoc);
            Assert.True(element != null, $"Failed to resolve anchor {kvp.Key}");
            Assert.Equal(target.Unid, (string?)element!.Attribute(PtOpenXml.Unid));
        }
    }

    [Theory]
    [InlineData("HC001-5DayTourPlanTemplate.docx")]
    public void MD004_AnchorsSurviveRoundTrip(string fixtureName)
    {
        var doc = LoadFixture(fixtureName);
        var first = WmlToMarkdownConverter.Convert(doc, new WmlToMarkdownConverterSettings());

        // Round-trip the in-memory bytes through OpenXmlMemoryStreamDocument.
        WmlDocument roundTripped;
        using (var sm = new OpenXmlMemoryStreamDocument(doc))
        {
            using (var w = sm.GetWordprocessingDocument())
            {
                w.MainDocumentPart!.Document.Save();
            }
            roundTripped = sm.GetModifiedWmlDocument();
        }

        var second = WmlToMarkdownConverter.Convert(roundTripped, new WmlToMarkdownConverterSettings());
        Assert.Equal(
            first.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray(),
            second.AnchorIndex.Keys.OrderBy(k => k, StringComparer.Ordinal).ToArray());
    }
}
