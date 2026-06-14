#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// FOLLOW-ON B: per-cell table composition in the consolidate engine. Today multi-reviewer edits to the
/// SAME base table route to a whole-block conflict (BaseWins keeps the base table; disjoint cell edits are
/// NOT composed). These tests pin the NEW behavior: DISJOINT cross-reviewer table-cell edits COMPOSE inline
/// (Alice edits cell(0,0), Bob edits cell(1,1) → both land, attributed); only SAME-cell edits by ≥2
/// reviewers → recorded conflict per policy. STOP boundaries (MovedRow / column-count / cell-tcPr changes)
/// fall back to the whole-table block conflict.
/// </summary>
public class IrCompositeTableTests
{
    // ------------------------------------------------------------------ fixtures

    private static string Cell(string text) =>
        $"<w:tc><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:tc>";

    private static string Row(params string[] cells) => $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid/>{string.Concat(rows)}</w:tbl>";

    private static string Lead => "<w:p><w:r><w:t>lead</w:t></w:r></w:p>";

    /// <summary>A 2x2 table base; one paragraph above so the body has stable surrounding structure.</summary>
    private static WmlDocument Base2x2() => IrTestDocuments.FromBodyXml(
        Lead +
        Table(
            Row(Cell("a one"), Cell("b two")),
            Row(Cell("c three"), Cell("d four"))));

    private static WmlDocument Variant2x2(string c00, string c01, string c10, string c11) =>
        IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell(c00), Cell(c01)),
                Row(Cell(c10), Cell(c11))));

    private static WmlDocument Consolidate(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.Consolidate(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static IReadOnlyList<DocxDiffConflict> Conflicts(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConflicts(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static IReadOnlyList<DocxDiffConsolidatedRevision> Revisions(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConsolidatedRevisions(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static string Json(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConsolidatedEditScriptJson(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static string Accept(WmlDocument merged) =>
        Docs.PlainTextWithTables(RevisionAccepter.AcceptRevisions(merged));

    private static string Reject(WmlDocument merged) =>
        Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged));

    // ------------------------------------------------------------------ 1. disjoint cells compose

    /// <summary>
    /// #1 — Disjoint cells compose. 2x2 base, Alice edits cell(0,0), Bob edits cell(1,1) → both land
    /// authored, Conflicts == 0. (RED before per-cell composition: whole-table conflict / one dropped.)
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Disjoint_cells_compose_both_land_no_conflict(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a ALICE", "b two", "c three", "d four"); // cell(0,0)
        var bob = Variant2x2("a one", "b two", "c three", "d BOB");      // cell(1,1) — disjoint

        // No conflict: disjoint cell edits compose.
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        // Accept yields BOTH edits regardless of policy (disjoint → policy-invariant).
        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("a ALICE", accept);
        Assert.Contains("d BOB", accept);
        Assert.Contains("b two", accept);
        Assert.Contains("c three", accept);

        // reject ≡ base.
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 2. consensus cell

    /// <summary>#2 — Consensus cell: both edit cell(0,0) identically → one cell, no conflict.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Consensus_cell_no_conflict(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a SAME", "b two", "c three", "d four");
        var bob = Variant2x2("a SAME", "b two", "c three", "d four"); // identical edit

        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("a SAME", accept);
        // The consensus edit appears once (no duplicated "SAME").
        Assert.Equal(1, CountOccurrences(accept, "SAME"));

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 3. same-cell different-words conflict

    /// <summary>
    /// #3 — Same-cell different-words conflict. Alice "the SLOW fox" / Bob "the QUICK fox" in cell(0,0) →
    /// a token conflict recorded at the cell-paragraph anchor; resolved per policy.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_cell_different_words_conflict(ConflictResolution policy)
    {
        var baseDoc = Variant2x2("the fox", "b two", "c three", "d four");
        var alice = Variant2x2("the SLOW fox", "b two", "c three", "d four");
        var bob = Variant2x2("the QUICK fox", "b two", "c three", "d four");

        var conflicts = Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.NotEmpty(conflicts);
        // The conflict competitors carry both reviewers' competing inserted words.
        var competitorTexts = conflicts.SelectMany(c => c.Competitors.Select(x => x.ResultText)).ToList();
        Assert.Contains(competitorTexts, t => t.Contains("SLOW"));
        Assert.Contains(competitorTexts, t => t.Contains("QUICK"));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        switch (policy)
        {
            case ConflictResolution.BaseWins:
                Assert.DoesNotContain("SLOW", accept);
                Assert.DoesNotContain("QUICK", accept);
                break;
            case ConflictResolution.FirstReviewerWins:
                Assert.Contains("SLOW", accept);
                Assert.DoesNotContain("QUICK", accept);
                break;
            case ConflictResolution.StackAll:
                Assert.Contains("SLOW", accept);
                Assert.Contains("QUICK", accept);
                break;
        }

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 4. same-cell disjoint words (recursion)

    /// <summary>
    /// #4 — Same-cell disjoint words: both reviewers edit DIFFERENT words of one cell paragraph → compose
    /// inside the cell (recursion proof: token composition runs at cell-paragraph granularity).
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_cell_disjoint_words_compose_inside_cell(ConflictResolution policy)
    {
        var baseDoc = Variant2x2("alpha beta gamma", "b two", "c three", "d four");
        var alice = Variant2x2("ALPHA beta gamma", "b two", "c three", "d four"); // first word
        var bob = Variant2x2("alpha beta GAMMA", "b two", "c three", "d four");   // last word

        // Disjoint words inside the SAME cell paragraph compose: no conflict.
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("ALPHA", accept);
        Assert.Contains("GAMMA", accept);
        Assert.Contains("beta", accept);

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 5. disjoint InsertRow by two reviewers

    /// <summary>#5 — Disjoint InsertRow by two reviewers → both rows appear, no conflict.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Disjoint_inserted_rows_both_appear(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("a one"), Cell("b two")),
                Row(Cell("c three"), Cell("d four")),
                Row(Cell("e ALICE"), Cell("f alice")))); // appends a row
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("a one"), Cell("b two")),
                Row(Cell("c three"), Cell("d four")),
                Row(Cell("g BOB"), Cell("h bob")))); // appends a different row

        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("e ALICE", accept);
        Assert.Contains("g BOB", accept);

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 6. DeleteRow (A) vs ModifyRow other row (B)

    /// <summary>#6 — DeleteRow (Alice deletes row 1) vs ModifyRow different row (Bob edits row 0) → both land.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void DeleteRow_vs_modify_other_row_both_land(ConflictResolution policy)
    {
        // 3-row base so a clean row delete leaves a stable spine.
        var baseDoc = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r1 a"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b"))));
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // deletes row 1
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a EDIT"), Cell("r0 b")),
                Row(Cell("r1 a"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // edits row 0 cell(0,0)

        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("EDIT", accept);          // Bob's row-0 edit landed
        Assert.DoesNotContain("r1 a", accept);    // Alice's row-1 delete landed (row gone on accept)

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 7. DeleteRow vs ModifyRow SAME row

    /// <summary>#7 — DeleteRow vs ModifyRow SAME row → row conflict, policy-resolved, reject ≡ base.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void DeleteRow_vs_modify_same_row_conflict(ConflictResolution policy)
    {
        var baseDoc = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r1 a"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b"))));
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // deletes row 1
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r1 a EDIT"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // edits row 1 (the one Alice deletes)

        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 8. MovedRow fallback (STOP boundary)

    /// <summary>
    /// #8 — MovedRow present → block-level conflict fallback (STOP boundary). When any reviewer's table diff
    /// has a MovedRow, the per-cell compose bails to the existing whole-table block conflict. No silent loss;
    /// reject ≡ base.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void MovedRow_falls_back_to_block_conflict(ConflictResolution policy)
    {
        // A 3-row table with UNIQUE row hashes so the row aligner anchors them; Alice reorders rows
        // (a MovedRow), Bob edits a cell — two reviewers touch the same base table.
        var baseDoc = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("alpha row")),
                Row(Cell("bravo row")),
                Row(Cell("charlie row"))));
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("bravo row")),
                Row(Cell("alpha row")),   // alpha moved down (a row move)
                Row(Cell("charlie row"))));
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("alpha row")),
                Row(Cell("bravo row EDIT")),
                Row(Cell("charlie row"))));

        // Falls back to a recorded whole-table conflict — no silent loss.
        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 9. column-count change fallback (gate)

    /// <summary>
    /// #9 — Column add/remove → block-level conflict fallback (gate test). When a reviewer adds a column
    /// (changing the cell count of a row) while another edits a different cell, the positional per-cell render
    /// (which clones the base row's cell skeleton, count-stable in v1) cannot compose the column change — so it
    /// bails to a recorded whole-table block conflict. No silent loss: a conflict is recorded under every
    /// policy, and BaseWins (which keeps the base table) rejects to base.
    /// <para>(A pure <c>w:tcPr</c> width change is NOT modeled by the IR — it leaves every hash identical, so
    /// the reviewer's table reads as EqualBlock (no touch) and composes trivially as single-source; only a
    /// genuine cell-count change is a detectable STOP boundary. reject ≡ base is NOT asserted for
    /// FirstReviewerWins/StackAll here: those emit the reviewer's column-changed table, and a 2-way column
    /// ADD is not yet marked as a tracked insertion — a pre-existing whole-table-renderer limitation,
    /// independent of this follow-on. The fallback's job is no-silent-loss, which holds.)</para>
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Column_count_change_falls_back_to_block_conflict(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("a one"), Cell("b two"), Cell("NEW col")), // adds a 3rd column to row 0
                Row(Cell("c three"), Cell("d four"))));
        var bob = Variant2x2("a one", "b two", "c three", "d BOB"); // edits cell(1,1) text

        // Column add → fall back to recorded whole-table conflict (no silent loss).
        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        // BaseWins keeps the base table → reject ≡ base (the fallback's reject contract on the base-kept path).
        if (policy == ConflictResolution.BaseWins)
        {
            var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
            Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
        }
    }

    // ------------------------------------------------------------------ 10. all 3 policies over same-cell conflict

    /// <summary>#10 — All three policies over a same-cell conflict (accept text + competitors).</summary>
    [Fact]
    public void Same_cell_conflict_all_policies()
    {
        var baseDoc = Variant2x2("base word", "b two", "c three", "d four");
        var alice = Variant2x2("ALICE word", "b two", "c three", "d four");
        var bob = Variant2x2("BOB word", "b two", "c three", "d four");

        // BaseWins: neither competitor's word in accept.
        var bw = Accept(Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob)));
        Assert.DoesNotContain("ALICE", bw);
        Assert.DoesNotContain("BOB", bw);

        // FirstReviewerWins: Alice's word, not Bob's.
        var fw = Accept(Consolidate(baseDoc, ConflictResolution.FirstReviewerWins, ("Alice", alice), ("Bob", bob)));
        Assert.Contains("ALICE", fw);
        Assert.DoesNotContain("BOB", fw);

        // StackAll: both.
        var sa = Accept(Consolidate(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob)));
        Assert.Contains("ALICE", sa);
        Assert.Contains("BOB", sa);

        // Competitors recorded with both authors.
        var conflicts = Conflicts(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));
        var authors = conflicts.SelectMany(c => c.Competitors.Select(x => x.Author)).Distinct().ToList();
        Assert.Contains("Alice", authors);
        Assert.Contains("Bob", authors);
    }

    // ------------------------------------------------------------------ 11. reject ≡ base all policies

    /// <summary>
    /// #11 — reject ≡ base for all three policies (markup → RejectRevisions → StructuralBody == base),
    /// across disjoint + same-cell + delete-row scenarios.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Reject_equals_base_all_policies(ConflictResolution policy)
    {
        var baseDoc = Base2x2();

        var disjoint = Consolidate(baseDoc, policy,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(disjoint));

        var sameCell = Consolidate(baseDoc, policy,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a BOB", "b two", "c three", "d four")));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(sameCell));
    }

    // ------------------------------------------------------------------ 12. renderer single w:tbl, authored cells

    /// <summary>
    /// #12 — Renderer: composed table = single w:tbl, per-cell w:ins/w:del authored to right reviewers, base
    /// cells verbatim; accept yields all disjoint edits.
    /// </summary>
    [Fact]
    public void Renderer_composed_table_single_tbl_authored_cells()
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a ALICE", "b two", "c three", "d four");
        var bob = Variant2x2("a one", "b two", "c three", "d BOB");

        var merged = Consolidate(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));
        var xml = Docs.MainPartXml(merged);

        // Exactly one table.
        Assert.Equal(1, CountOccurrences(xml, "<w:tbl>"));
        // Both reviewers' edits are present as tracked insertions attributed to them.
        Assert.Contains("w:author=\"Alice\"", xml);
        Assert.Contains("w:author=\"Bob\"", xml);
        // Untouched cells survive verbatim.
        Assert.Contains("b two", xml);
        Assert.Contains("c three", xml);

        var accept = Accept(merged);
        Assert.Contains("a ALICE", accept);
        Assert.Contains("d BOB", accept);
    }

    // ------------------------------------------------------------------ 13. revisions per reviewer / per conflict

    /// <summary>
    /// #13 — Revisions: disjoint-cell → revisions per reviewer; same-cell conflict → competing revisions
    /// linked to the cell conflict id.
    /// </summary>
    [Fact]
    public void Revisions_disjoint_per_reviewer_and_samecell_linked_to_conflict()
    {
        var baseDoc = Base2x2();

        // Disjoint: a revision for each reviewer, attributed.
        var disjointRevs = Revisions(baseDoc, ConflictResolution.StackAll,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")));
        Assert.Contains(disjointRevs, r => r.Author == "Alice");
        Assert.Contains(disjointRevs, r => r.Author == "Bob");

        // Same-cell conflict: competing revisions carry a conflict id matching a recorded conflict.
        var scBase = Variant2x2("base word", "b two", "c three", "d four");
        var scAlice = Variant2x2("ALICE word", "b two", "c three", "d four");
        var scBob = Variant2x2("BOB word", "b two", "c three", "d four");
        var conflicts = Conflicts(scBase, ConflictResolution.StackAll, ("Alice", scAlice), ("Bob", scBob));
        var revs = Revisions(scBase, ConflictResolution.StackAll, ("Alice", scAlice), ("Bob", scBob));
        var conflictIds = conflicts.Select(c => c.Id).ToHashSet();
        Assert.Contains(revs, r => r.ConflictId is { } id && conflictIds.Contains(id));
    }

    // ------------------------------------------------------------------ 14. JSON merged tableDiff + authoredRows

    /// <summary>
    /// #14 — JSON: composed table op serializes merged tableDiff + additive authoredRows; deterministic.
    /// </summary>
    [Fact]
    public void Json_composed_table_serializes_tableDiff_and_authoredRows_deterministic()
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a ALICE", "b two", "c three", "d four");
        var bob = Variant2x2("a one", "b two", "c three", "d BOB");

        var json1 = Json(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));
        var json2 = Json(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));

        // Deterministic.
        Assert.Equal(json1, json2);
        // The merged tableDiff is present (apply/json truth).
        Assert.Contains("\"tableDiff\"", json1);
        Assert.Contains("\"rowOps\"", json1);
        // The additive authoredRows attribution view is present.
        Assert.Contains("\"authoredRows\"", json1);
    }

    // ------------------------------------------------------------------ 15. verifier over disjoint + conflict

    /// <summary>
    /// #15 — Verifier proves the table compose: the composite apply-reconstruction (cell ops applied to the
    /// base table) matches the rendered accepted body, for disjoint and same-cell conflict cases, all policies.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Verifier_proves_table_compose_all_policies(ConflictResolution policy)
    {
        var baseDoc = Base2x2();

        var disjointMerged = Consolidate(baseDoc, policy,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")));
        IrCompositeVerifier.Verify(
            baseDoc,
            new (string, WmlDocument)[]
            {
                ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
                ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")),
            },
            policy,
            Docs.AcceptStructuralBody(disjointMerged));

        var scBase = Variant2x2("base word", "b two", "c three", "d four");
        var scAlice = Variant2x2("ALICE word", "b two", "c three", "d four");
        var scBob = Variant2x2("BOB word", "b two", "c three", "d four");
        var scMerged = Consolidate(scBase, policy, ("Alice", scAlice), ("Bob", scBob));
        IrCompositeVerifier.Verify(
            scBase,
            new (string, WmlDocument)[] { ("Alice", scAlice), ("Bob", scBob) },
            policy,
            Docs.AcceptStructuralBody(scMerged));
    }

    // ------------------------------------------------------------------ helpers

    private static int CountOccurrences(string haystack, string needle)
    {
        int count = 0, idx = 0;
        while ((idx = haystack.IndexOf(needle, idx, System.StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += needle.Length;
        }
        return count;
    }
}
