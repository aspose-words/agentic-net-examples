using System;
using Aspose.Words;

public class TrackChangesErrorHandling
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial text (no revision will be created for this).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions with a specific author.
        doc.StartTrackRevisions("Alice", DateTime.Now);

        // Add a new paragraph – this will be recorded as an insertion revision.
        builder.Writeln("Inserted paragraph.");

        // Remove the first paragraph to create a deletion revision.
        // The paragraph to delete is the first one (index 0).
        Paragraph paraToDelete = doc.FirstSection.Body.Paragraphs[0];
        paraToDelete.Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // At this point the document has two revisions:
        // 0 – Deletion of the original paragraph.
        // 1 – Insertion of the new paragraph.
        Console.WriteLine($"Total revisions after modifications: {doc.Revisions.Count}");

        // Reject the deletion revision (index 0).
        Revision deletionRevision = doc.Revisions[0];
        deletionRevision.Reject();
        Console.WriteLine("Deletion revision rejected.");

        // Attempt to accept the same revision again.
        // Since the revision has been removed from the collection, this should raise an exception.
        try
        {
            deletionRevision.Accept();
            Console.WriteLine("Unexpectedly accepted a rejected revision.");
        }
        catch (Exception ex)
        {
            // Handle the error gracefully.
            Console.WriteLine($"Error while accepting a rejected revision: {ex.Message}");
        }

        // Verify the remaining revisions count.
        Console.WriteLine($"Revisions count after handling: {doc.Revisions.Count}");

        // Save the resulting document.
        doc.Save("TrackChangesErrorHandling.docx");
    }
}
