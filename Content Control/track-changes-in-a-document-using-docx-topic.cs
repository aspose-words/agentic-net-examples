using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class TrackChangesDemo
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add initial content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Write("This text is not a revision. ");

        // Verify that no revisions exist yet.
        Console.WriteLine($"Revisions before tracking: {doc.Revisions.Count}");

        // Start tracking revisions with an author name.
        doc.StartTrackRevisions("John Doe");

        // Add content that will be recorded as an insertion revision.
        builder.Write("First tracked insertion. ");

        // Display revision count after the first tracked change.
        Console.WriteLine($"Revisions after first insertion: {doc.Revisions.Count}");
        Console.WriteLine($"First revision type: {doc.Revisions[0].RevisionType}");
        Console.WriteLine($"First revision author: {doc.Revisions[0].Author}");

        // Stop tracking revisions – subsequent edits will not be recorded.
        doc.StopTrackRevisions();

        // Add content that will NOT be recorded as a revision.
        builder.Write("This text is not tracked. ");

        Console.WriteLine($"Revisions after stopping tracking: {doc.Revisions.Count}");

        // Start tracking again, this time providing a custom date/time for the revisions.
        doc.StartTrackRevisions("John Doe", DateTime.MinValue);
        builder.Write("Second tracked insertion with custom date. ");

        // Verify the second revision's date.
        Console.WriteLine($"Revisions after second insertion: {doc.Revisions.Count}");
        Console.WriteLine($"Second revision date: {doc.Revisions[1].DateTime}");

        // Demonstrate accepting a specific revision (the first one).
        doc.Revisions[0].Accept();

        // After acceptance, the revision collection is reduced.
        Console.WriteLine($"Revisions after accepting first revision: {doc.Revisions.Count}");

        // Accept all remaining revisions.
        doc.AcceptAllRevisions();

        // Save the resulting document.
        doc.Save("TrackedChangesDemo.docx");
    }
}
