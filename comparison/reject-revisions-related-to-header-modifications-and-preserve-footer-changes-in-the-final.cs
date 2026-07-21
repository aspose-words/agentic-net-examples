using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create the original document with a header, footer and body text.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Header
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Original Header");

        // Footer
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        builder.Writeln("Original Footer");

        // Body
        builder.MoveToDocumentEnd();
        builder.Writeln("Original body content.");

        // Create the revised document with changed header and footer.
        Document revised = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revised);

        // Header (edited)
        revBuilder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        revBuilder.Writeln("Edited Header");

        // Footer (edited)
        revBuilder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
        revBuilder.Writeln("Edited Footer");

        // Body (unchanged)
        revBuilder.MoveToDocumentEnd();
        revBuilder.Writeln("Original body content.");

        // Compare the documents. The original will contain revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Collect revisions to avoid modifying the collection while iterating.
        List<Revision> revisions = original.Revisions.Cast<Revision>().ToList();

        // Reject revisions that belong to headers, keep footer revisions.
        foreach (Revision rev in revisions)
        {
            HeaderFooter ancestor = rev.ParentNode?.GetAncestor(NodeType.HeaderFooter) as HeaderFooter;
            if (ancestor != null && ancestor.HeaderFooterType != HeaderFooterType.FooterPrimary)
            {
                rev.Reject();
            }
        }

        // Accept all remaining revisions (footer changes).
        original.AcceptAllRevisions();

        // Save the final document.
        original.Save("Result.docx");
    }
}
