#nullable enable
using System.Linq;
using Docxodus;
using Xunit;
namespace Docxodus.Tests.Ir.Diff;
public class CompositeFuzzTests
{
    [Theory]
    [InlineData(3)] [InlineData(4)] [InlineData(5)]
    public void Composite_round_trips_reject_equals_base(int reviewerCount)
    {
        for (int seed = 0; seed < 50; seed++)
        {
            var fc = DiffFuzzer.GenerateComposite(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var reviewers = fc.Reviewers
                .Select(r => new DocxDiffReviewer { Document = new WmlDocument("r.docx", r.Doc), Author = r.Author })
                .ToList();
            var merged = DocxDiff.Consolidate(baseDoc, reviewers);
            Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        }
    }

    [Theory]
    [InlineData(3)] [InlineData(4)]
    public void Composite_apply_verifier_holds(int reviewerCount)
    {
        for (int seed = 0; seed < 30; seed++)
        {
            var fc = DiffFuzzer.GenerateComposite(seed, reviewerCount);
            var baseDoc = new WmlDocument("b.docx", fc.Base);
            var revs = fc.Reviewers.Select(r => (r.Author, (WmlDocument)new WmlDocument("r.docx", r.Doc))).ToList();
            var dd = revs.Select(r => new DocxDiffReviewer { Document = r.Item2, Author = r.Author }).ToList();
            var merged = DocxDiff.Consolidate(baseDoc, dd);
            var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
            IrCompositeVerifier.Verify(baseDoc, revs, ConflictResolution.BaseWins, acceptedText);
        }
    }
}
