using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content without tracking.
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with author "Alice".
        doc.StartTrackRevisions("Alice");

        // Make some changes to generate revisions.
        builder.Writeln("First revision paragraph.");

        // Delete a run to create a deletion revision.
        // This operation is performed while tracking is active.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after the first tracking session.");

        // Accept all revisions.
        doc.AcceptAllRevisions();

        // Verify that all revisions have been accepted.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were not fully accepted.");

        // Save the document after accepting revisions (optional intermediate file).
        doc.Save("AfterAccept.docx");

        // Re‑enable tracking for subsequent changes with a different author.
        doc.StartTrackRevisions("Bob");

        // Add more content to generate new revisions.
        builder.Writeln("Second revision paragraph.");

        // Stop tracking again.
        doc.StopTrackRevisions();

        // Verify that new revisions exist.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after the second tracking session.");

        // Save the final document.
        doc.Save("FinalDocument.docx");
    }
}
