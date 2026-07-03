#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Header/footer scope diffing — reader section-reference capture, builder pairing, revisions,
/// and edit-script JSON (spec: 2026-07-03-docxdiff-header-footer-diff-design.md). The markup
/// renderer's story tests live in <see cref="IrMarkupRendererTests"/>.
/// </summary>
public class IrHeaderFooterDiffTests
{
    // ------------------------------------------------------------------ reader: section references

    [Fact]
    public void Reader_records_section_references_in_document_order()
    {
        // Two sections. Section 0: default header A + first header B. Section 1: default header A
        // again (multi-referenced part). Footer C on section 1 only.
        var doc = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "sec0 body" },
                    Headers: new[] { ("default", "rIdA"), ("first", "rIdB") }),
                new HeaderFooterFixtures.Section(
                    new[] { "sec1 body" },
                    Headers: new[] { ("default", "rIdA") },
                    Footers: new[] { ("default", "rIdC") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Header A" },
                ["rIdB"] = new[] { "Header B" },
            },
            footerParts: new Dictionary<string, string[]> { ["rIdC"] = new[] { "Footer C" } },
            titlePg: true);

        var ir = IrReader.Read(doc);

        var a = Assert.Single(ir.Headers, h => h.Scope.Blocks.Any(b => BlockText(b).Contains("Header A")));
        Assert.Equal(new[] { (0, IrHeaderFooterKind.Default), (1, IrHeaderFooterKind.Default) },
            a.References.Select(r => (r.SectionIndex, r.Kind)).ToArray());
        Assert.Equal(IrHeaderFooterKind.Default, a.Kind);

        var b = Assert.Single(ir.Headers, h => h.Scope.Blocks.Any(bl => BlockText(bl).Contains("Header B")));
        Assert.Equal(new[] { (0, IrHeaderFooterKind.First) },
            b.References.Select(r => (r.SectionIndex, r.Kind)).ToArray());
        Assert.Equal(IrHeaderFooterKind.First, b.Kind);

        var c = Assert.Single(ir.Footers);
        Assert.Equal(new[] { (1, IrHeaderFooterKind.Default) },
            c.References.Select(r => (r.SectionIndex, r.Kind)).ToArray());
    }

    [Fact]
    public void Reader_unreferenced_part_has_empty_references()
    {
        // One referenced header + one orphan header part no sectPr cites.
        var doc = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "body" }, Headers: new[] { ("default", "rIdA") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdA"] = new[] { "Referenced" },
                ["rIdOrphan"] = new[] { "Orphan" },
            });

        var ir = IrReader.Read(doc);

        var orphan = Assert.Single(ir.Headers, h => h.Scope.Blocks.Any(b => BlockText(b).Contains("Orphan")));
        Assert.Empty(orphan.References);
        Assert.Equal(IrHeaderFooterKind.Default, orphan.Kind);
    }

    private static string BlockText(IrBlock block) =>
        block is IrParagraph p
            ? string.Concat(p.Inlines.OfType<IrTextRun>().Select(r => r.Text))
            : string.Empty;
}
