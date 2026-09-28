using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with a header, footer and body text.
        Document original = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(original);

        // Header
        builderOrig.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builderOrig.Writeln("Original Header");

        // Footer
        builderOrig.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builderOrig.Writeln("Original Footer");

        // Body
        builderOrig.MoveToDocumentEnd();
        builderOrig.Writeln("Body text.");

        // Save the original for reference (optional).
        original.Save("original.docx");

        // Create the revised document with modified header and footer.
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);

        // Modified Header
        builderRev.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builderRev.Writeln("Revised Header");

        // Modified Footer
        builderRev.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builderRev.Writeln("Revised Footer");

        // Body (unchanged)
        builderRev.MoveToDocumentEnd();
        builderRev.Writeln("Body text.");

        // Save the revised for reference (optional).
        revised.Save("revised.docx");

        // Compare the documents. The comparison result is stored in 'original'.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated during comparison.");

        // Copy revisions to an array because the collection changes when accepting/rejecting.
        Revision[] revisions = original.Revisions.Cast<Revision>().ToArray();

        foreach (Revision rev in revisions)
        {
            // The node that changed is typically a Paragraph inside a HeaderFooter.
            HeaderFooter? parentHeaderFooter = rev.ParentNode?.ParentNode as HeaderFooter;

            if (parentHeaderFooter != null)
            {
                // Determine if the revision is in a header.
                if (parentHeaderFooter.HeaderFooterType == HeaderFooterType.HeaderPrimary ||
                    parentHeaderFooter.HeaderFooterType == HeaderFooterType.HeaderFirst ||
                    parentHeaderFooter.HeaderFooterType == HeaderFooterType.HeaderEven)
                {
                    rev.Reject(); // Discard header changes.
                }
                else
                {
                    rev.Accept(); // Keep footer changes.
                }
            }
            else
            {
                // For any other revision types (e.g., body text), accept them.
                rev.Accept();
            }
        }

        // After processing, the document should contain only the accepted footer revision.
        // Save the final document.
        original.Save("final_output.docx");
    }
}
