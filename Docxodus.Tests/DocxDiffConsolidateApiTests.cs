#nullable enable
using System.Linq;
using Docxodus;
using Xunit;
namespace Docxodus.Tests;
public class DocxDiffConsolidateApiTests
{
    [Fact]
    public void Settings_default_policy_is_base_wins()
    {
        var s = new DocxDiffConsolidateSettings();
        Assert.Equal(ConflictResolution.BaseWins, s.ConflictResolution);
        Assert.NotNull(s.Diff);
    }

    [Fact]
    public void Reviewer_holds_document_and_author()
    {
        var r = new DocxDiffReviewer { Author = "Bob" };
        Assert.Equal("Bob", r.Author);
    }

    [Fact]
    public void Consolidate_two_reviewers_round_trips()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("alpha one", "beta two", "gamma three");
        var r1 = Docxodus.Tests.Ir.Diff.Docs.Para("alpha one EDITED", "beta two", "gamma three");
        var r2 = Docxodus.Tests.Ir.Diff.Docs.Para("alpha one", "beta two", "gamma three EDITED");
        var merged = DocxDiff.Consolidate(b, new[]
        {
            new DocxDiffReviewer { Document = r1, Author = "Bob" },
            new DocxDiffReviewer { Document = r2, Author = "Fred" },
        });
        Assert.Equal(Docxodus.Tests.Ir.Diff.Docs.PlainText(b),
            Docxodus.Tests.Ir.Diff.Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docxodus.Tests.Ir.Diff.Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("alpha one EDITED", accepted);
        Assert.Contains("gamma three EDITED", accepted);
    }

    [Fact]
    public void GetConflicts_reports_overlapping_edit_with_competitors()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("the quick brown fox");
        var r1 = Docxodus.Tests.Ir.Diff.Docs.Para("the SLOW brown fox");
        var r2 = Docxodus.Tests.Ir.Diff.Docs.Para("the FAST brown fox");
        var conflicts = DocxDiff.GetConflicts(b, new[]
        {
            new DocxDiffReviewer { Document = r1, Author = "Bob" },
            new DocxDiffReviewer { Document = r2, Author = "Fred" },
        });
        Assert.NotEmpty(conflicts);
        Assert.Contains(conflicts.SelectMany(c => c.Competitors), x => x.Author == "Bob");
        Assert.Contains(conflicts.SelectMany(c => c.Competitors), x => x.Author == "Fred");
    }

    [Fact]
    public void Consolidate_zero_reviewers_returns_base()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("x");
        var merged = DocxDiff.Consolidate(b, System.Array.Empty<DocxDiffReviewer>());
        Assert.Equal(Docxodus.Tests.Ir.Diff.Docs.PlainText(b), Docxodus.Tests.Ir.Diff.Docs.PlainText(merged));
    }
}
