using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content without tracking (these are not revisions).
        builder.Writeln("Paragraph 1: Original content.");
        builder.Writeln("Paragraph 2: Original content.");

        // ---------- First tracking session ----------
        // Enable tracking to generate revisions.
        doc.StartTrackRevisions("Author1");

        // Make some changes that will be recorded as revisions.
        builder.Writeln("Paragraph 3: Added while tracking.");
        // Delete a run to create a deletion revision.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        if (firstParagraph.Runs.Count > 0)
        {
            firstParagraph.Runs[0].Remove();
        }

        // Stop tracking; further edits will not be revisions.
        doc.StopTrackRevisions();

        // Accept all revisions generated in the first session.
        doc.AcceptAllRevisions();

        // At this point the document should have no pending revisions.
        if (doc.HasRevisions)
        {
            throw new InvalidOperationException("Revisions were not fully accepted.");
        }

        // ---------- Second tracking session ----------
        // Re‑enable tracking to capture new changes separately.
        doc.StartTrackRevisions("Author2");

        // Add more content; these will appear as new revisions.
        builder.Writeln("Paragraph 4: Added in second tracking session.");
        // Modify formatting (note: formatting changes are not tracked as revisions by Aspose.Words).

        // Stop tracking again.
        doc.StopTrackRevisions();

        // Save the resulting document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TrackChangesDemo.docx");
        doc.Save(outputPath);
    }
}
