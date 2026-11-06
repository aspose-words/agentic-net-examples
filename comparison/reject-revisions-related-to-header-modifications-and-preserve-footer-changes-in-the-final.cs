using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create the original document with header, footer and body.
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

        // Create the revised document with modified header and footer.
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);

        // Header (modified)
        builderRev.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builderRev.Writeln("Revised Header");

        // Footer (modified)
        builderRev.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builderRev.Writeln("Revised Footer");

        // Body (unchanged)
        builderRev.MoveToDocumentEnd();
        builderRev.Writeln("Body text.");

        // Compare the documents to generate revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Process revisions: reject those in headers, accept all others.
        List<Revision> revisions = new List<Revision>(original.Revisions);
        foreach (Revision rev in revisions)
        {
            // Determine if the revision is inside a header.
            HeaderFooter? headerFooter = rev.ParentNode?.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            if (headerFooter != null && headerFooter.IsHeader)
            {
                rev.Reject(); // Discard header changes.
            }
            else
            {
                rev.Accept(); // Keep footer (and any other) changes.
            }
        }

        // Save the final document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "final.docx");
        original.Save(outputPath);
    }
}
