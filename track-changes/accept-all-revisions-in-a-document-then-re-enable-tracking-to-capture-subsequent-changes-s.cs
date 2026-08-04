using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content without tracking – this will not create revisions.
        builder.Writeln("Original paragraph. ");

        // ---------- First tracking session ----------
        // Enable tracking with the first author.
        doc.StartTrackRevisions("Author1");

        // Make changes that will be recorded as revisions.
        builder.Writeln("First revision paragraph. ");
        builder.Writeln("Another line added while tracking. ");

        // Stop tracking – further edits will not be revisions.
        doc.StopTrackRevisions();

        // At this point the document should contain revisions.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Accept all revisions, removing them from the collection.
        doc.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (doc.HasRevisions || doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // ---------- Second tracking session ----------
        // Re‑enable tracking with a different author to capture new changes separately.
        doc.StartTrackRevisions("Author2");

        // Add more content – these will appear as new revisions.
        builder.Writeln("Second revision paragraph after acceptance. ");
        builder.Writeln("Additional text for the second tracking session. ");

        // Stop tracking again.
        doc.StopTrackRevisions();

        // Save the final document to disk.
        doc.Save("Output.docx");
    }
}
