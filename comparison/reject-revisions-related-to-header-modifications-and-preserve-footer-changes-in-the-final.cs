using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the original document with a header, a footer and some body text.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original Header");

        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Original Footer");

        builder.MoveToDocumentEnd();
        builder.Writeln("Body content.");

        // Create the revised document where the header and footer are changed.
        Document revised = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(revised);

        builder2.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder2.Writeln("Edited Header");

        builder2.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder2.Writeln("Edited Footer");

        builder2.MoveToDocumentEnd();
        builder2.Writeln("Body content.");

        // Compare the documents – revisions will be created in the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Reject revisions that belong to the header, keep (accept) all other revisions.
        var revisions = original.Revisions;
        var revisionList = revisions.Cast<Revision>().ToList(); // snapshot to avoid collection modification issues

        foreach (Revision rev in revisionList)
        {
            // Determine if the revision is inside a header.
            HeaderFooter headerFooter = rev.ParentNode?.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            bool isHeaderRevision = headerFooter != null && headerFooter.HeaderFooterType == HeaderFooterType.HeaderPrimary;

            if (isHeaderRevision)
                rev.Reject();   // Discard header changes.
            else
                rev.Accept();   // Preserve footer (and other) changes.
        }

        // Save the final document where header changes are rejected and footer changes are kept.
        original.Save("Result.docx");
    }
}
