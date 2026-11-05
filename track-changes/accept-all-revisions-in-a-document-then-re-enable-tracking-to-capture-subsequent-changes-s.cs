using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // 1. Create a new blank document and add initial content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Save the original document (no revisions).
        doc.Save("Original.docx");

        // 2. Enable track changes to generate revisions.
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Make some modifications while tracking is enabled.
        builder.MoveToDocumentEnd();
        builder.Writeln("This paragraph was added while tracking changes.");
        builder.MoveToDocumentStart();
        builder.Write("Modified ");

        // Stop tracking changes.
        doc.StopTrackRevisions();

        // Save the document that now contains revisions.
        doc.Save("WithRevisions.docx");

        // 3. Accept all revisions that were generated.
        doc.AcceptAllRevisions();

        // Save the document after accepting revisions.
        doc.Save("Accepted.docx");

        // 4. Re‑enable tracking to capture subsequent changes separately.
        doc.StartTrackRevisions("Author2", DateTime.Now);

        // Add more content while tracking is active.
        builder.MoveToDocumentEnd();
        builder.Writeln("Another paragraph added after accepting previous revisions.");

        // Stop tracking again.
        doc.StopTrackRevisions();

        // Save the final document which contains the new set of revisions.
        doc.Save("Final.docx");
    }
}
