#nullable enable
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeJsonTests
{
    [Fact]
    public void Consolidated_edit_script_json_has_author_and_conflicts()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox"); var r2 = Docs.Para("the FAST brown fox");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains("\"author\"", json);
        Assert.Contains("\"conflicts\"", json);
        using var _ = System.Text.Json.JsonDocument.Parse(json); // valid JSON
    }

    [Fact]
    public void Consolidated_edit_script_json_carries_source_reviewer_and_anchors()
    {
        var b = Docs.Para("alpha one", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "gamma three");
        var r2 = Docs.Para("alpha one", "gamma three EDITED");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains("\"sourceReviewer\"", json);
        Assert.Contains("\"operations\"", json);
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("conflicts", out _));
        Assert.True(doc.RootElement.TryGetProperty("operations", out _));
    }

    [Fact]
    public void Zero_reviewers_json_is_empty_operations_and_conflicts()
    {
        var b = Docs.Para("alpha one");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, System.Array.Empty<DocxDiffReviewer>());
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        Assert.Empty(doc.RootElement.GetProperty("operations").EnumerateArray());
        Assert.Empty(doc.RootElement.GetProperty("conflicts").EnumerateArray());
    }

    /// <summary>
    /// CONTRACT: a composite over a reviewer MOVE lowers the move SOURCE to a DeleteBlock and the move
    /// DESTINATION to an InsertBlock (see <see cref="Docxodus.Ir.Diff.IrCompositeMerger.LowerStructuralOps"/>).
    /// The lowered move-source DeleteBlock RETAINS MoveGroupId/IsMoveSource INTERNALLY (the merger's
    /// contested-relocation detection reads them), but the documented <c>IrEditOp</c> field-presence contract
    /// (IrEditScript.cs) says a DeleteBlock/InsertBlock carries NULL move fields — so the EMITTED, serialized
    /// op must NOT leak <c>moveGroupId</c>/<c>isMoveSource</c>. The edit-script-as-data is the public
    /// differentiator; a machine consumer must see a contract-clean delete/insert.
    /// </summary>
    [Fact]
    public void Consolidated_json_delete_insert_ops_carry_no_move_fields()
    {
        // A clean single-reviewer relocation of a ≥4-word paragraph so move detection fires and the move is
        // lowered to a source DeleteBlock + a destination InsertBlock in the composite.
        const string p1 = "First paragraph alpha bravo";
        const string p2 = "Second paragraph charlie delta";
        const string p3 = "Third paragraph echo foxtrot";
        const string p4 = "Fourth paragraph golf hotel";
        var b = Docs.Para(p1, p2, p3, p4);
        var alice = Docs.Para(p1, p3, p4, p2); // P2 relocated to the end (a move → lowered to del + ins)

        var json = DocxDiff.GetConsolidatedEditScriptJson(
            b, new[] { new DocxDiffReviewer { Document = alice, Author = "Alice" } });

        using var doc = System.Text.Json.JsonDocument.Parse(json);
        foreach (var op in doc.RootElement.GetProperty("operations").EnumerateArray())
        {
            var kind = op.GetProperty("kind").GetString();
            if (kind is "DeleteBlock" or "InsertBlock")
            {
                Assert.False(op.TryGetProperty("moveGroupId", out _),
                    $"{kind} leaked moveGroupId into the public edit-script JSON (IrEditOp contract violation).");
                Assert.False(op.TryGetProperty("isMoveSource", out _),
                    $"{kind} leaked isMoveSource into the public edit-script JSON (IrEditOp contract violation).");
            }
        }

        // Sanity: the move WAS lowered (a DeleteBlock for the move source is present), so the assertion above
        // is not vacuously true.
        Assert.Contains(doc.RootElement.GetProperty("operations").EnumerateArray(),
            op => op.GetProperty("kind").GetString() == "DeleteBlock");
    }
}
