#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// N-way merger that consolidates several pairwise <see cref="IrEditScript"/>s — each computed
/// against the SAME base document, so they share one base anchor space — into a single
/// <see cref="IrCompositeScript"/>.
/// <para>The merge walks the base document in block order. For each base block it gathers the
/// reviewers' ops anchored there (via <see cref="GroupByBaseAnchor"/>) and dispatches to
/// <c>MergeOneBaseBlock</c>; right-only inserts are slotted relative to the preceding base anchor
/// (via <c>GroupInsertsByPrecedingAnchor</c> + <c>EmitInsertsAt</c>). Conflicts — a base block
/// edited differently by 2+ reviewers — are resolved by the supplied
/// <see cref="Docxodus.ConflictResolution"/> policy and recorded in
/// <see cref="IrCompositeScript.Conflicts"/>.</para>
/// <para>See <c>docs/architecture/ir_diff_engine.md</c> for the diff pipeline this composes over.</para>
/// </summary>
internal static class IrCompositeMerger
{
    public static IrCompositeScript Merge(
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings)
    {
        var scripts = reviewers.Select(r => IrEditScriptBuilder.Build(baseIr, r.Ir, settings)).ToList();
        var baseOrder = BaseBlockAnchors(baseIr);
        var byBase = GroupByBaseAnchor(scripts);
        var insertsAfter = GroupInsertsByPrecedingAnchor(scripts, baseOrder);

        var ops = new List<IrCompositeOp>();
        var conflicts = new List<IrConflict>();
        var nextConflictId = 1;
        EmitInsertsAt(insertsAfter, "", reviewers, ops);
        foreach (var anchor in baseOrder)
        {
            MergeOneBaseBlock(anchor, byBase, reviewers, baseIr, policy, settings, ops, conflicts, ref nextConflictId);
            EmitInsertsAt(insertsAfter, anchor, reviewers, ops);
        }
        return new IrCompositeScript(IrNodeList.From(ops), IrNodeList.From(conflicts));
    }

    /// <summary>
    /// Index every op that carries a non-null <see cref="IrEditOp.LeftAnchor"/> (i.e. it touches a
    /// base block) by that anchor string, tagged with its reviewer index in <paramref name="scripts"/>.
    /// </summary>
    internal static Dictionary<string, List<(int Reviewer, IrEditOp Op)>> GroupByBaseAnchor(
        IReadOnlyList<IrEditScript> scripts)
    {
        var byBase = new Dictionary<string, List<(int Reviewer, IrEditOp Op)>>();
        for (int reviewer = 0; reviewer < scripts.Count; reviewer++)
        {
            foreach (var op in scripts[reviewer].Operations)
            {
                if (op.LeftAnchor is not { } anchor)
                    continue;
                if (!byBase.TryGetValue(anchor, out var list))
                    byBase[anchor] = list = new List<(int, IrEditOp)>();
                list.Add((reviewer, op));
            }
        }
        return byBase;
    }

    /// <summary>Base body blocks in document order, as <c>kind:scope:unid</c> anchor strings.</summary>
    internal static List<string> BaseBlockAnchors(IrDocument baseIr) =>
        baseIr.Body.Blocks.Select(b => b.Anchor.ToString()).ToList();

    // ---- Task 1.3: exact sub-block (token-span) composition ----

    /// <summary>
    /// Compose N reviewers' per-base-paragraph token diffs — all expressed over the SAME base token
    /// stream <c>[0, baseTokenCount)</c> — into a single authored token-op list plus any conflicts
    /// (spec §6 step 4). Walks base positions <c>pos = 0..baseTokenCount</c> inclusive; at each pos it
    /// (1) resolves the reviewers' INSERTs anchored at pos (consensus when ≤1 distinct inserted text,
    /// otherwise a conflict resolved by <paramref name="policy"/>), then (2) for the base token
    /// <c>[pos, pos+1)</c> resolves the reviewers' DELETEs (consensus when 0/1 deleter or all deleters
    /// delete identically, otherwise a conflict). Conflict ids ascend from <paramref name="conflictIdSeed"/>.
    /// </summary>
    /// <remarks>
    /// INVARIANT (span totality): the emitted ops' non-insert left spans tile <c>[0, baseTokenCount)</c>
    /// exactly once — every base token becomes exactly one Equal or one Delete; inserts carry empty
    /// left spans. Verified by <see cref="AssertTilesBase"/> under <c>Debug.Assert</c>.
    /// </remarks>
    internal static List<IrAuthoredTokenOp> ComposeTokenSpans(
        int baseTokenCount,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, int conflictIdSeed,
        out List<IrConflict> conflicts)
    {
        var result = new List<IrAuthoredTokenOp>();
        conflicts = new List<IrConflict>();
        int nextConflictId = conflictIdSeed;

        // Replacement convention (matches IrTokenDiffer): the differ emits a replacement as a Delete
        // followed by an Insert whose LeftStart == Delete.LeftEnd (NOT Delete.LeftStart). The
        // immediately-following Insert in the same reviewer's Ops is that Delete's replacement text.
        // Such replacement-Inserts are owned by the delete path (ComposeBaseTokenAt), keyed to the
        // DELETED base span — so we exclude them here from the standalone insert-at-pos handling to
        // avoid double-counting one reviewer's "replace base[a,b) with text" as both a delete-conflict
        // at [a,b) AND an insert-conflict at [b,b). We identify the set by reference identity.
        var replacementInserts = CollectReplacementInserts(reviewerDiffs);

        for (int pos = 0; pos <= baseTokenCount; pos++)
        {
            ComposeInsertsAt(pos, reviewerDiffs, replacementInserts, policy, baseAnchor, ref nextConflictId, result, conflicts);
            if (pos < baseTokenCount)
                ComposeBaseTokenAt(pos, reviewerDiffs, policy, baseAnchor, ref nextConflictId, result, conflicts);
        }

        System.Diagnostics.Debug.Assert(AssertTilesBase(result, baseTokenCount),
            "ComposeTokenSpans: emitted left spans must tile [0, baseTokenCount) exactly once.");
        return result;
    }

    /// <summary>
    /// The set (by reference identity) of every Insert op that is the replacement half of a
    /// Delete+Insert replacement pair, across all reviewers. A replacement Insert is the op that
    /// immediately follows a Delete in the same reviewer's <see cref="IrTokenDiff.Ops"/> and whose
    /// <c>LeftStart == Delete.LeftEnd</c> (the differ's replacement anchoring). These are handled by the
    /// delete path and must NOT separately re-trigger an insert conflict.
    /// </summary>
    private static HashSet<IrTokenOp> CollectReplacementInserts(
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs)
    {
        var set = new HashSet<IrTokenOp>(ReferenceEqualityComparer.Instance);
        foreach (var rd in reviewerDiffs)
        {
            var ops = rd.Diff.Ops;
            for (int k = 0; k + 1 < ops.Count; k++)
            {
                if (ops[k].Kind == IrTokenOpKind.Delete &&
                    ops[k + 1].Kind == IrTokenOpKind.Insert &&
                    ops[k + 1].LeftStart == ops[k].LeftEnd)
                {
                    set.Add(ops[k + 1]);
                }
            }
        }
        return set;
    }

    /// <summary>Step 1: resolve every reviewer INSERT anchored at base position <paramref name="pos"/>.</summary>
    private static void ComposeInsertsAt(
        int pos,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        HashSet<IrTokenOp> replacementInserts,
        Docxodus.ConflictResolution policy, string baseAnchor, ref int nextConflictId,
        List<IrAuthoredTokenOp> result, List<IrConflict> conflicts)
    {
        // Each entry: the inserting reviewer, in reviewer (list) order, with the inserted text.
        // Replacement-Inserts (the Insert half of a Delete+Insert pair) are EXCLUDED — they are owned by
        // the delete path keyed to the deleted base span, so they must not re-trigger an insert conflict.
        var inserters = new List<(int ReviewerIndex, string Author, IrTokenOp Op, string Text)>();
        for (int i = 0; i < reviewerDiffs.Count; i++)
        {
            var rd = reviewerDiffs[i];
            foreach (var op in rd.Diff.Ops)
            {
                if (op.Kind == IrTokenOpKind.Insert && op.LeftStart == pos && !replacementInserts.Contains(op))
                    inserters.Add((rd.Reviewer, rd.Author, op, RightText(op, rd.RightTokens)));
            }
        }
        if (inserters.Count == 0)
            return;

        // Group by inserted right text, preserving first-appearance order (reviewer order).
        var distinctTexts = new List<string>();
        foreach (var ins in inserters)
            if (!distinctTexts.Contains(ins.Text))
                distinctTexts.Add(ins.Text);

        if (distinctTexts.Count <= 1)
        {
            // Consensus (single reviewer, or multiple reviewers inserting the SAME text): one authored
            // Insert per distinct text; the first reviewer in that group is the author/source.
            foreach (var text in distinctTexts)
            {
                var first = inserters.First(x => x.Text == text);
                result.Add(new IrAuthoredTokenOp(first.Op, first.Author, first.ReviewerIndex));
            }
            return;
        }

        // ≥2 distinct text groups → CONFLICT: one competitor per inserting reviewer.
        var competitors = inserters
            .Select(x => new IrConflictCompetitor(x.Author, x.Text))
            .ToList();
        conflicts.Add(new IrConflict(
            nextConflictId++, baseAnchor, pos, pos, policy, IrNodeList.From(competitors)));

        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                // Emit NOTHING for the inserts at pos.
                break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
            {
                var first = inserters.OrderBy(x => x.ReviewerIndex).First();
                result.Add(new IrAuthoredTokenOp(first.Op, first.Author, first.ReviewerIndex));
                break;
            }
            case Docxodus.ConflictResolution.StackAll:
                foreach (var ins in inserters.OrderBy(x => x.ReviewerIndex))
                    result.Add(new IrAuthoredTokenOp(ins.Op, ins.Author, ins.ReviewerIndex));
                break;
        }
    }

    /// <summary>Step 2: resolve the fate of base token <c>[pos, pos+1)</c> across reviewer DELETEs.</summary>
    /// <remarks>
    /// <para><b>Per-base-token tiling vs multi-token deletes.</b> A reviewer's Delete may span several
    /// base tokens (e.g. the differ deletes a word AND its trailing separator as <c>Delete(2,4|2,2)</c>).
    /// This step runs once per base token and emits a SINGLE-token Delete <c>[pos, pos+1)</c> for each
    /// deleted base token — never the deleter's whole multi-token span — so the emitted left spans tile
    /// <c>[0, baseTokenCount)</c> exactly once (the totality invariant). A multi-token delete therefore
    /// surfaces as several adjacent single-token Delete ops, authored to the same reviewer.</para>
    /// <para><b>Replacement insert anchoring.</b> A replacement's Insert is logically anchored at
    /// <c>Delete.LeftEnd</c> (one past the last deleted token). To emit it EXACTLY once for a multi-token
    /// replacement, we emit it only when resolving the LAST covered base token (<c>pos == Delete.LeftEnd - 1</c>).
    /// The insert path (<see cref="ComposeInsertsAt"/>) excludes replacement-Inserts, so this is the only
    /// site that emits the replacement text.</para>
    /// </remarks>
    private static void ComposeBaseTokenAt(
        int pos,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, ref int nextConflictId,
        List<IrAuthoredTokenOp> result, List<IrConflict> conflicts)
    {
        // Each deleter, in reviewer order, with the covering Delete op and its replacement text ("" for
        // a pure delete). A replacement is the Insert op IMMEDIATELY FOLLOWING the covering Delete in the
        // same reviewer's Ops (the differ anchors it at Delete.LeftEnd, NOT Delete.LeftStart). Built in
        // reviewer (list) order so deleters[0] / deleters.First() == the lowest reviewer index (M1).
        var deleters = new List<(int ReviewerIndex, string Author, IrTokenOp DeleteOp, string Replacement)>();
        for (int i = 0; i < reviewerDiffs.Count; i++)
        {
            var rd = reviewerDiffs[i];
            if (!TryFindCoveringDelete(rd.Diff.Ops, pos, out var delOp, out var replacement, rd.RightTokens))
                continue;
            deleters.Add((rd.Reviewer, rd.Author, delOp!, replacement));
        }

        if (deleters.Count == 0)
        {
            // Base survives.
            result.Add(BaseEqual(pos));
            return;
        }

        // "Identical" deletion: all deleters have byte-equal replacement text (all pure deletes, or
        // all delete-and-replace with the same inserted text). Pure-delete vs delete-with-replacement
        // are NOT identical (one has "" replacement, the other does not).
        bool allIdentical = deleters.All(d => d.Replacement == deleters[0].Replacement);

        if (deleters.Count == 1 || allIdentical)
        {
            // Consensus: delete THIS base token once, authored to the first deleter (lowest reviewer
            // index). If this is a delete-with-replacement, also emit its replacement insert — but only
            // when resolving the deleter's LAST covered token, so a multi-token replacement emits the
            // text exactly once. The insert path excludes replacement-Inserts, so this is the only site.
            var first = deleters[0];
            result.Add(SingleTokenDelete(pos, first.Author, first.ReviewerIndex));
            EmitReplacementInsertIfLast(pos, first.DeleteOp, first.Author, first.ReviewerIndex, reviewerDiffs, result);
            return;
        }

        // ≥2 deleters that are NOT identical → CONFLICT (recorded per base token of the contested span).
        var competitors = deleters
            .Select(d => new IrConflictCompetitor(d.Author, d.Replacement))
            .ToList();
        conflicts.Add(new IrConflict(
            nextConflictId++, baseAnchor, pos, pos + 1, policy, IrNodeList.From(competitors)));

        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                // Base survives.
                result.Add(BaseEqual(pos));
                break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
            {
                // First deleter (lowest reviewer index): delete this base token AND emit its replacement
                // insert (once, at the deleter's last covered token) so the winning replacement lands.
                var first = deleters[0];
                result.Add(SingleTokenDelete(pos, first.Author, first.ReviewerIndex));
                EmitReplacementInsertIfLast(pos, first.DeleteOp, first.Author, first.ReviewerIndex, reviewerDiffs, result);
                break;
            }
            case Docxodus.ConflictResolution.StackAll:
            {
                // A base token can only be deleted once: delete THIS base token once (authored to the
                // first deleter). The competing REPLACEMENT inserts are excluded from the insert path, so
                // stack them here (in reviewer order), each emitted once at its own deleter's last token.
                var first = deleters[0];
                result.Add(SingleTokenDelete(pos, first.Author, first.ReviewerIndex));
                foreach (var d in deleters)
                    EmitReplacementInsertIfLast(pos, d.DeleteOp, d.Author, d.ReviewerIndex, reviewerDiffs, result);
                break;
            }
        }
    }

    /// <summary>
    /// Emit <paramref name="reviewerIndex"/>'s replacement Insert for <paramref name="delOp"/> — but only
    /// when <paramref name="pos"/> is the deleter's LAST covered base token (<c>pos == delOp.LeftEnd - 1</c>),
    /// so a multi-token replacement emits its inserted text exactly once.
    /// </summary>
    private static void EmitReplacementInsertIfLast(
        int pos, IrTokenOp delOp, string author, int reviewerIndex,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        List<IrAuthoredTokenOp> result)
    {
        if (pos != delOp.LeftEnd - 1)
            return;
        var ins = FindReplacementInsert(reviewerDiffs, reviewerIndex, delOp);
        if (ins is not null)
            result.Add(new IrAuthoredTokenOp(ins, author, reviewerIndex));
    }

    /// <summary>
    /// The replacement Insert op for <paramref name="delOp"/> in reviewer <paramref name="reviewerIndex"/>:
    /// the op immediately following <paramref name="delOp"/> in that reviewer's Ops when it is an Insert
    /// anchored at <c>delOp.LeftEnd</c>, else null. Mirrors <see cref="TryFindCoveringDelete"/>'s pairing
    /// rule; used to RE-emit a replacement's text since the insert path excludes replacement-Inserts.
    /// </summary>
    private static IrTokenOp? FindReplacementInsert(
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        int reviewerIndex, IrTokenOp delOp)
    {
        foreach (var rd in reviewerDiffs)
        {
            if (rd.Reviewer != reviewerIndex)
                continue;
            var ops = rd.Diff.Ops;
            for (int k = 0; k + 1 < ops.Count; k++)
            {
                if (ReferenceEquals(ops[k], delOp) &&
                    ops[k + 1].Kind == IrTokenOpKind.Insert &&
                    ops[k + 1].LeftStart == delOp.LeftEnd)
                {
                    return ops[k + 1];
                }
            }
            return null;
        }
        return null;
    }

    /// <summary>A base-sourced Equal op covering base token <paramref name="pos"/> (Author "", reviewer -1).</summary>
    private static IrAuthoredTokenOp BaseEqual(int pos) =>
        new(new IrTokenOp(IrTokenOpKind.Equal, pos, pos + 1, pos, pos + 1), "", -1);

    /// <summary>
    /// A single-base-token Delete <c>[pos, pos+1)</c> authored to <paramref name="author"/> /
    /// <paramref name="reviewerIndex"/>. Always single-token (never a reviewer's whole multi-token Delete
    /// span) so the per-base-token loop preserves the <c>[0, baseTokenCount)</c> tiling invariant.
    /// </summary>
    private static IrAuthoredTokenOp SingleTokenDelete(int pos, string author, int reviewerIndex) =>
        new(new IrTokenOp(IrTokenOpKind.Delete, pos, pos + 1, pos, pos), author, reviewerIndex);

    /// <summary>
    /// Find the Delete op in <paramref name="ops"/> whose left span covers base token
    /// <paramref name="pos"/> (<c>LeftStart &lt;= pos &lt; LeftEnd</c>), and resolve its replacement
    /// text. The replacement is the op IMMEDIATELY FOLLOWING that Delete in <paramref name="ops"/> when
    /// that op is an Insert anchored at <c>Delete.LeftEnd</c> (the differ's Delete→Insert replacement
    /// convention) — concatenated right text — else "" (pure delete). Op-list adjacency is the
    /// authoritative pairing rule; the <c>LeftStart == Delete.LeftEnd</c> check additionally guards
    /// against an unrelated standalone Insert that merely happens to follow the Delete in Ops.
    /// </summary>
    private static bool TryFindCoveringDelete(
        IrNodeList<IrTokenOp> ops, int pos, out IrTokenOp? delOp, out string replacement,
        IReadOnlyList<IrDiffToken> rightTokens)
    {
        for (int k = 0; k < ops.Count; k++)
        {
            var op = ops[k];
            if (op.Kind != IrTokenOpKind.Delete || op.LeftStart > pos || pos >= op.LeftEnd)
                continue;

            delOp = op;
            replacement =
                k + 1 < ops.Count &&
                ops[k + 1].Kind == IrTokenOpKind.Insert &&
                ops[k + 1].LeftStart == op.LeftEnd
                    ? RightText(ops[k + 1], rightTokens)
                    : "";
            return true;
        }

        delOp = null;
        replacement = "";
        return false;
    }

    /// <summary>Concatenate the raw <see cref="IrDiffToken.Text"/> of <paramref name="op"/>'s right span.</summary>
    private static string RightText(IrTokenOp op, IReadOnlyList<IrDiffToken> rightTokens)
    {
        if (op.RightLength == 0)
            return "";
        var sb = new System.Text.StringBuilder();
        for (int i = op.RightStart; i < op.RightEnd; i++)
            sb.Append(rightTokens[i].Text);
        return sb.ToString();
    }

    /// <summary>
    /// Debug-only totality check: the concatenation of non-insert (Equal/Delete) left spans, in op
    /// order, equals <c>[0, baseTokenCount)</c> exactly once ascending.
    /// </summary>
    private static bool AssertTilesBase(List<IrAuthoredTokenOp> ops, int baseTokenCount)
    {
        int expected = 0;
        foreach (var authored in ops)
        {
            var op = authored.Op;
            if (op.Kind == IrTokenOpKind.Insert)
                continue;
            if (op.LeftStart != expected)
                return false;
            expected = op.LeftEnd;
        }
        return expected == baseTokenCount;
    }

    // ---- Filled by Task 1.4 ----

    internal static Dictionary<string, List<(int Reviewer, IrEditOp Op)>> GroupInsertsByPrecedingAnchor(
        IReadOnlyList<IrEditScript> scripts, List<string> baseOrder) => throw new System.NotImplementedException();

    private static void EmitInsertsAt(
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> insertsAfter,
        string anchor,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        List<IrCompositeOp> ops) => throw new System.NotImplementedException();

    private static void MergeOneBaseBlock(
        string anchor,
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> byBase,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings,
        List<IrCompositeOp> ops,
        List<IrConflict> conflicts,
        ref int nextConflictId) => throw new System.NotImplementedException();
}
