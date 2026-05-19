using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text – this is not a revision.
        builder.Write("Initial text. ");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe", DateTime.Now);
        builder.Write("Tracked change. ");

        // At this point one revision (the insertion) should exist.
        if (doc.Revisions.Count != 1)
            throw new InvalidOperationException("Expected 1 revision after tracking started.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Record the revision count after stopping tracking.
        int revisionsBefore = doc.Revisions.Count;

        // Make additional edits – these should NOT be recorded as revisions.
        builder.Write("Untracked change. ");

        // Verify that no new revisions were added.
        int revisionsAfter = doc.Revisions.Count;
        if (revisionsAfter != revisionsBefore)
            throw new InvalidOperationException("New revisions were recorded after tracking was stopped.");

        // Additionally, confirm that the last run is not marked as an insert revision.
        Run lastRun = (Run)doc.FirstSection.Body.LastParagraph.Runs[doc.FirstSection.Body.LastParagraph.Runs.Count - 1];
        if (lastRun.IsInsertRevision)
            throw new InvalidOperationException("The last run is incorrectly marked as an insert revision.");

        // Save the document to verify the result manually if needed.
        doc.Save("TrackChangesResult.docx");
    }
}
