using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("Paragraph 1: Original text.");
        builder.Writeln("Paragraph 2: Original text.");

        // 2. Enable tracking to generate first set of revisions.
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Make some changes while tracking is enabled.
        // Insert a new paragraph.
        builder.Writeln("Paragraph 3: Inserted while tracking.");
        // Modify existing paragraph.
        builder.MoveToDocumentStart();
        builder.Writeln("Paragraph 0: Inserted at the beginning.");

        // Stop tracking.
        doc.StopTrackRevisions();

        // At this point the document contains revisions.
        Console.WriteLine($"Revisions after first edit: {doc.Revisions.Count}");

        // 3. Accept all existing revisions.
        doc.AcceptAllRevisions();

        // Verify that revisions have been cleared.
        Console.WriteLine($"Revisions after acceptance: {doc.Revisions.Count}");

        // 4. Re‑enable tracking to capture subsequent changes separately.
        doc.StartTrackRevisions("Author2", DateTime.Now);

        // Make additional changes.
        // Delete the second paragraph (original Paragraph 1).
        Paragraph paraToDelete = (Paragraph)doc.GetChild(NodeType.Paragraph, 1, true);
        paraToDelete.Remove();

        // Change formatting of the first paragraph.
        Paragraph firstPara = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        firstPara.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Stop tracking again.
        doc.StopTrackRevisions();

        // Final revision count should reflect only the second set of changes.
        Console.WriteLine($"Revisions after second edit: {doc.Revisions.Count}");

        // 5. Save the final document.
        doc.Save("TrackChangesDemo.docx");
    }
}
