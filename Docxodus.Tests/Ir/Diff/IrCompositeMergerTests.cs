#nullable enable
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeMergerTests
{
    internal static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    [Fact]
    public void GroupByBaseAnchor_colocates_per_reviewer_ops()
    {
        var baseDoc = Docs.Para("alpha one", "beta two");
        var r1 = Docs.Para("alpha one EDITED", "beta two");
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var s1 = IrEditScriptBuilder.Build(baseIr, IrReader.Read(r1, ReadOpts), diff);
        var grouped = IrCompositeMerger.GroupByBaseAnchor(new[] { s1 });
        Assert.Contains(grouped.Values, list => list.Any(x => x.Op.Kind == IrEditOpKind.ModifyBlock));
    }

    [Fact]
    public void Disjoint_block_edits_from_two_reviewers_both_appear_attributed()
    {
        var b = Docs.Para("alpha one", "beta two", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "beta two", "gamma three");
        var r2 = Docs.Para("alpha one", "beta two", "gamma three EDITED");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        var mods = s.Operations.Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock).ToList();
        Assert.Equal(2, mods.Count);
        Assert.Contains(mods, m => m.Author == "Bob" || (m.AuthoredTokens?.Any(a => a.Author == "Bob") ?? false));
        Assert.Contains(mods, m => m.Author == "Fred" || (m.AuthoredTokens?.Any(a => a.Author == "Fred") ?? false));
        Assert.Empty(s.Conflicts);
    }

    [Fact]
    public void Identical_edit_by_both_reviewers_is_consensus_not_conflict()
    {
        var b = Docs.Para("alpha one", "beta two");
        var same = Docs.Para("alpha one EDITED", "beta two");
        var s = MergeOf(b, ("Bob", same), ("Fred", same));
        Assert.Empty(s.Conflicts);
        Assert.Single(s.Operations.Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock));
    }

    [Fact]
    public void Two_reviewers_editing_same_words_differently_conflict()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox");
        var r2 = Docs.Para("the FAST brown fox");
        Assert.NotEmpty(MergeOf(b, ("Bob", r1), ("Fred", r2)).Conflicts);
    }

    [Fact]
    public void Two_reviewers_editing_different_words_of_one_paragraph_compose_without_conflict()
    {
        var b = Docs.Para("the quick brown fox jumps");
        var r1 = Docs.Para("the SLOW brown fox jumps");
        var r2 = Docs.Para("the quick brown fox LEAPS");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        Assert.Empty(s.Conflicts);
        var mod = Assert.Single(s.Operations.Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock));
        Assert.NotNull(mod.AuthoredTokens);
        Assert.Contains(mod.AuthoredTokens!, a => a.Author == "Bob");
        Assert.Contains(mod.AuthoredTokens!, a => a.Author == "Fred");
    }

    [Fact]
    public void Both_reviewers_inserting_paragraphs_both_appear_no_conflict()
    {
        var b = Docs.Para("alpha", "omega");
        var r1 = Docs.Para("alpha", "bob inserted line", "omega");
        var r2 = Docs.Para("alpha", "fred inserted line", "omega");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        Assert.Equal(2, s.Operations.Count(o => o.Op.Kind == IrEditOpKind.InsertBlock));
        Assert.Empty(s.Conflicts);
    }

    [Fact]
    public void Delete_by_one_modify_by_other_is_block_conflict()
    {
        var b = Docs.Para("alpha", "beta two three", "omega");
        var deleter = Docs.Para("alpha", "omega");
        var editor = Docs.Para("alpha", "beta two three EDITED", "omega");
        Assert.NotEmpty(MergeOf(b, ("Bob", deleter), ("Fred", editor)).Conflicts);
    }

    [Fact]
    public void Both_reviewers_deleting_same_block_is_consensus_not_conflict()
    {
        var b = Docs.Para("alpha", "beta to delete", "omega");
        var r1 = Docs.Para("alpha", "omega");
        var r2 = Docs.Para("alpha", "omega");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        Assert.Empty(s.Conflicts);
        Assert.Single(s.Operations.Where(o => o.Op.Kind == IrEditOpKind.DeleteBlock));
    }

    // Helper reused by later tasks: merge base + reviewers into an IrCompositeScript.
    internal static IrCompositeScript MergeOf(WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
        => MergeOf(ConflictResolution.BaseWins, baseDoc, reviewers);

    internal static IrCompositeScript MergeOf(ConflictResolution policy, WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
    {
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var revs = reviewers.Select(r => (r.Author, IrReader.Read(r.Doc, ReadOpts))).ToList();
        return IrCompositeMerger.Merge(baseIr, revs, policy, diff);
    }
}
