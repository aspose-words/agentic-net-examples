using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text – this should not create any revisions.
        builder.Write("Original text. ");

        // Verify that there are no revisions yet.
        if (doc.Revisions.Count != 0)
            throw new Exception("Expected zero revisions before tracking starts.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Author1", DateTime.Now);

        // Add text while tracking is enabled – this should create a revision.
        builder.Write("First revision. ");

        // Verify that exactly one revision was recorded.
        if (doc.Revisions.Count != 1)
            throw new Exception("Expected one revision after first tracked edit.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Add more text after tracking has been stopped – this should NOT create a new revision.
        builder.Write("After stop tracking. ");

        // Verify that the revision count has not increased.
        if (doc.Revisions.Count != 1)
            throw new Exception("No new revisions should be recorded after stopping tracking.");

        // Additionally, confirm that the last run is not marked as an insertion revision.
        var runs = doc.FirstSection.Body.FirstParagraph.Runs;
        var lastRun = runs[runs.Count - 1];
        if (lastRun.IsInsertRevision)
            throw new Exception("The text added after stopping tracking should not be marked as a revision.");

        // Save the document to the local file system.
        doc.Save("TrackChangesDemo.docx");
    }
}
