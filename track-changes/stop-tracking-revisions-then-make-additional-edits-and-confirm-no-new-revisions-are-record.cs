using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content (not tracked).
        builder.Write("Initial content. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make a change that will be recorded as a revision.
        builder.Write("First tracked change. ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Capture the revision count after stopping tracking.
        int revisionsAfterStop = doc.Revisions.Count;

        // Make another change after tracking has been stopped.
        builder.Write("Second change after stop. ");

        // Verify that no new revisions have been added.
        if (doc.Revisions.Count != revisionsAfterStop)
            throw new InvalidOperationException("New revisions were recorded after StopTrackRevisions.");

        // Save the document (optional, demonstrates that the document can be persisted).
        doc.Save("TrackChangesDemo.docx");
    }
}
