using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text without tracking – this should not create revisions.
        builder.Write("Initial text. ");

        // Verify that no revisions exist yet.
        if (doc.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions should be zero before tracking starts.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe");

        // Write text while tracking – this should create a revision.
        builder.Write("Tracked change 1. ");

        // Verify that one revision was recorded.
        if (doc.Revisions.Count != 1)
            throw new InvalidOperationException("Exactly one revision should exist after first tracked edit.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Write additional text after stopping tracking – this must NOT create new revisions.
        builder.Write("Untracked change after stop. ");

        // Verify that the revision count has not increased.
        if (doc.Revisions.Count != 1)
            throw new InvalidOperationException("No new revisions should be recorded after stopping tracking.");

        // Save the document to verify the result (optional for the task).
        doc.Save("TrackedChangesExample.docx");

        // Output confirmation.
        Console.WriteLine("Revisions count after operations: " + doc.Revisions.Count);
        Console.WriteLine("Example completed successfully.");
    }
}
