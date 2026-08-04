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

        // Write some initial content (not tracked).
        builder.Write("Original text. ");

        // Verify that no revisions exist yet.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Document should have zero revisions before tracking.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make an edit that will be recorded as a revision.
        builder.Write("First revision. ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Capture the revision count after the tracked edit.
        int revisionCountAfterFirstEdit = doc.Revisions.Count;

        // Ensure that the edit created exactly one revision.
        if (revisionCountAfterFirstEdit != 1)
            throw new InvalidOperationException("Expected exactly one revision after the first edit.");

        // Perform additional edits after tracking has been stopped.
        builder.Write("Edit after stopping tracking. ");

        // Verify that no new revisions were added.
        if (doc.Revisions.Count != revisionCountAfterFirstEdit)
            throw new InvalidOperationException("New revisions were recorded after StopTrackRevisions.");

        // Save the document to verify the result.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "TrackChangesOutput.docx");
        doc.Save(outputPath);
    }
}
