using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- First revision session ----------
        // Enable tracking with the first author.
        doc.StartTrackRevisions("Author1");

        // Make some changes that will be recorded as revisions.
        builder.Writeln("This is the first revision.");

        // Stop tracking so further changes are not recorded.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions were not created in the first session.");

        // Accept all revisions, clearing the revision collection.
        doc.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (doc.HasRevisions || doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // ---------- Second revision session ----------
        // Re‑enable tracking with a different author.
        doc.StartTrackRevisions("Author2");

        // Make additional changes that should be captured as new revisions.
        builder.Writeln("This is the second revision, captured after accepting the first.");

        // Stop tracking again.
        doc.StopTrackRevisions();

        // Verify that new revisions exist.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions were not created in the second session.");

        // Save the final document.
        doc.Save("Output.docx");
    }
}
