using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Initial content.");

        // Enable tracking and make the first revision.
        doc.StartTrackRevisions("Author1", DateTime.Now);
        builder.Writeln("First revision text.");
        doc.StopTrackRevisions();

        // Save the document that now contains revisions.
        doc.Save("WithRevisions.docx");

        // Accept all revisions that were made.
        doc.AcceptAllRevisions();

        // Save the document after accepting revisions.
        doc.Save("Accepted.docx");

        // Re‑enable tracking to capture subsequent changes separately.
        doc.StartTrackRevisions("Author2", DateTime.Now);
        builder.Writeln("Second revision after acceptance.");
        doc.StopTrackRevisions();

        // Save the final document.
        doc.Save("Final.docx");
    }
}
