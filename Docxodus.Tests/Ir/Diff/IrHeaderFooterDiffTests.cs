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

    // ------------------------------------------------------------------ builder: story pairing

    private static readonly IrDiffSettings Default = new();

    private static IrEditScript Build(WmlDocument left, WmlDocument right, IrDiffSettings? settings = null) =>
        IrEditScriptBuilder.Build(IrReader.Read(left), IrReader.Read(right), settings ?? Default);

    [Fact]
    public void Builder_matched_story_produces_token_diff()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "CONFIDENTIAL Draft 1" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "CONFIDENTIAL Draft 2" });

        var script = Build(left, right);

        Assert.NotNull(script.HeaderFooterOps);
        var diff = Assert.Single(script.HeaderFooterOps!);
        Assert.True(diff.IsHeader);
        Assert.Equal(IrHeaderFooterKind.Default, diff.Kind);
        Assert.Equal(0, diff.SectionIndex);
        Assert.Equal("hdr1", diff.ScopeName);
        Assert.Equal("hdr1", diff.LeftScopeName);
        Assert.NotNull(diff.LeftPartUri);
        Assert.NotNull(diff.RightPartUri);
        var op = Assert.Single(diff.Ops);
        Assert.Equal(IrEditOpKind.ModifyBlock, op.Kind);
        Assert.NotNull(op.TokenDiff);
        Assert.StartsWith("p:hdr1:", op.LeftAnchor);
        Assert.StartsWith("p:hdr1:", op.RightAnchor);
    }

    [Fact]
    public void Builder_unchanged_story_emits_no_header_footer_ops()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body one" }, headerParas: new[] { "Same header" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body two" }, headerParas: new[] { "Same header" });

        var script = Build(left, right);

        Assert.Null(script.HeaderFooterOps);
    }

    [Fact]
    public void Builder_inserted_story_is_all_insert_blocks()
    {
        // Right adds a FIRST-page header the left lacks.
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Running" });
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "Body" },
                    Headers: new[] { ("default", "rIdH1"), ("first", "rIdH2") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdH1"] = new[] { "Running" },
                ["rIdH2"] = new[] { "Cover page banner" },
            },
            titlePg: true);

        var script = Build(left, right);

        Assert.NotNull(script.HeaderFooterOps);
        var diff = Assert.Single(script.HeaderFooterOps!);
        Assert.True(diff.IsHeader);
        Assert.Equal(IrHeaderFooterKind.First, diff.Kind);
        Assert.Null(diff.LeftScopeName);
        Assert.Null(diff.LeftPartUri);
        Assert.NotNull(diff.RightPartUri);
        Assert.All(diff.Ops, o => Assert.Equal(IrEditOpKind.InsertBlock, o.Kind));
    }

    [Fact]
    public void Builder_deleted_story_is_all_delete_blocks()
    {
        var left = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Running" }, footerParas: new[] { "Old footer" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Running" });

        var script = Build(left, right);

        Assert.NotNull(script.HeaderFooterOps);
        var diff = Assert.Single(script.HeaderFooterOps!);
        Assert.False(diff.IsHeader);
        Assert.Equal("ftr1", diff.ScopeName);
        Assert.Equal("ftr1", diff.LeftScopeName);
        Assert.NotNull(diff.LeftPartUri);
        Assert.Null(diff.RightPartUri);
        Assert.All(diff.Ops, o => Assert.Equal(IrEditOpKind.DeleteBlock, o.Kind));
    }

    [Fact]
    public void Builder_orders_headers_before_footers()
    {
        var left = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Header v1" }, footerParas: new[] { "Footer v1" });
        var right = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Header v2" }, footerParas: new[] { "Footer v2" });

        var script = Build(left, right);

        Assert.NotNull(script.HeaderFooterOps);
        Assert.Equal(2, script.HeaderFooterOps!.Count);
        Assert.True(script.HeaderFooterOps[0].IsHeader);
        Assert.False(script.HeaderFooterOps[1].IsHeader);
    }

    [Fact]
    public void Builder_pairs_by_section_and_kind_not_scope_name()
    {
        // Left declares parts in order (first, default): the DEFAULT story is hdr2.
        // Right declares (default, first): the DEFAULT story is hdr1. Only the default story changes.
        var left = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "Body" },
                    Headers: new[] { ("first", "rIdF"), ("default", "rIdD") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdF"] = new[] { "Cover" },
                ["rIdD"] = new[] { "Running v1" },
            },
            titlePg: true);
        var right = HeaderFooterFixtures.Build(
            new[]
            {
                new HeaderFooterFixtures.Section(
                    new[] { "Body" },
                    Headers: new[] { ("default", "rIdD"), ("first", "rIdF") }),
            },
            headerParts: new Dictionary<string, string[]>
            {
                ["rIdD"] = new[] { "Running v2" },
                ["rIdF"] = new[] { "Cover" },
            },
            titlePg: true);

        var script = Build(left, right);

        Assert.NotNull(script.HeaderFooterOps);
        var diff = Assert.Single(script.HeaderFooterOps!);
        Assert.Equal(IrHeaderFooterKind.Default, diff.Kind);
        Assert.Equal("hdr2", diff.LeftScopeName); // left part order: first=hdr1, default=hdr2
        Assert.Equal("hdr1", diff.ScopeName);     // right part order: default=hdr1
    }

    [Fact]
    public void Builder_inherited_story_diffs_once()
    {
        // Two sections; section 1 has NO header references and inherits section 0's default story.
        HeaderFooterFixtures.Section[] Sections() => new[]
        {
            new HeaderFooterFixtures.Section(new[] { "sec0 body" }, Headers: new[] { ("default", "rIdA") }),
            new HeaderFooterFixtures.Section(new[] { "sec1 body" }),
        };
        var left = HeaderFooterFixtures.Build(Sections(),
            headerParts: new Dictionary<string, string[]> { ["rIdA"] = new[] { "Inherited v1" } });
        var right = HeaderFooterFixtures.Build(Sections(),
            headerParts: new Dictionary<string, string[]> { ["rIdA"] = new[] { "Inherited v2" } });

        var script = Build(left, right);

        Assert.NotNull(script.HeaderFooterOps);
        var diff = Assert.Single(script.HeaderFooterOps!);
        Assert.Equal(0, diff.SectionIndex);
    }

    [Fact]
    public void Builder_gate_off_restores_old_behavior()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Header v1" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body" }, headerParas: new[] { "Header v2" });

        var script = Build(left, right, new IrDiffSettings { CompareHeadersFooters = false });

        Assert.Null(script.HeaderFooterOps);
    }

    // ------------------------------------------------------------------ edit-script JSON

    [Fact]
    public void Json_round_trips_header_footer_ops()
    {
        var left = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Header v1" }, footerParas: new[] { "Same footer" });
        var right = HeaderFooterFixtures.Simple(
            new[] { "Body" }, headerParas: new[] { "Header v2" }, footerParas: new[] { "Same footer" });

        var script = Build(left, right);
        Assert.NotNull(script.HeaderFooterOps);

        var json = IrEditScriptJson.Write(script);
        Assert.Contains("\"headerFooterOps\"", json);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));
    }

    [Fact]
    public void Json_omits_header_footer_ops_when_no_story_changed()
    {
        var left = HeaderFooterFixtures.Simple(new[] { "Body one" }, headerParas: new[] { "Same" });
        var right = HeaderFooterFixtures.Simple(new[] { "Body two" }, headerParas: new[] { "Same" });

        var json = IrEditScriptJson.Write(Build(left, right));

        Assert.DoesNotContain("headerFooterOps", json);
    }

    [Fact]
    public void Public_settings_map_compare_headers_footers()
    {
        Assert.True(new DocxDiffSettings().CompareHeadersFooters);
        Assert.True(new DocxDiffSettings().ToIrDiffSettings().CompareHeadersFooters);
        Assert.False(new DocxDiffSettings { CompareHeadersFooters = false }
            .ToIrDiffSettings().CompareHeadersFooters);
    }
}
