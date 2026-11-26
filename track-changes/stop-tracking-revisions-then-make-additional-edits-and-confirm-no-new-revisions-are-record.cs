using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (not tracked).
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make a change while tracking is enabled – this will create a revision.
        builder.Writeln("Inserted while tracking.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Record the number of revisions after stopping tracking.
        int revisionsAfterStop = doc.Revisions.Count;

        // Make another change after tracking has been stopped – this should NOT create a revision.
        builder.Writeln("Inserted after tracking stopped.");

        // Verify that no new revisions were added.
        int revisionsAfterEdit = doc.Revisions.Count;
        if (revisionsAfterEdit != revisionsAfterStop)
        {
            throw new InvalidOperationException("A new revision was recorded after tracking was stopped.");
        }

        // Output the final revision count.
        Console.WriteLine($"Final revision count: {revisionsAfterEdit}");

        // Save the document to verify the result (optional for the task).
        doc.Save("TrackChangesDemo.docx");
    }
}
