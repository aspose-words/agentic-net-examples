using System;
using System.Collections.Generic;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with two sections.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Section 1
        builder.Writeln("Section 1 - Original text.");
        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        // Section 2
        builder.Writeln("Section 2 - Original text.");

        // Start tracking revisions.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Make changes in both sections to generate revisions.
        // Change in Section 1.
        builder.MoveToDocumentStart();
        builder.Writeln("Inserted line in Section 1.");

        // Change in Section 2.
        builder.MoveToDocumentEnd();
        builder.Writeln("Inserted line in Section 2.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        int totalRevisions = doc.Revisions.Count;
        if (totalRevisions == 0)
            throw new InvalidOperationException("No revisions were generated.");

        // Accept revisions only in the first section.
        List<Revision> revisionsToAccept = new List<Revision>();
        foreach (Revision rev in doc.Revisions)
        {
            Node sectionNode = rev.ParentNode.GetAncestor(NodeType.Section);
            if (sectionNode != null && doc.Sections.IndexOf((Section)sectionNode) == 0)
                revisionsToAccept.Add(rev);
        }

        foreach (Revision rev in revisionsToAccept)
            rev.Accept();

        // After acceptance, there should still be revisions (those from Section 2).
        int remainingRevisions = doc.Revisions.Count;
        if (remainingRevisions == 0)
            throw new InvalidOperationException("All revisions were accepted; expected some to remain.");

        // Save the resulting document.
        doc.Save("Result.docx");

        // Output counts for verification.
        Console.WriteLine($"Total revisions generated: {totalRevisions}");
        Console.WriteLine($"Revisions accepted in Section 1: {revisionsToAccept.Count}");
        Console.WriteLine($"Revisions remaining (Section 2): {remainingRevisions}");
    }
}
