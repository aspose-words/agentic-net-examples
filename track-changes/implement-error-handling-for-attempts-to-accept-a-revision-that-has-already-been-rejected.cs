using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write initial text (no revision yet).
        builder.Write("Original text. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Insert new text – this will be an insertion revision.
        builder.Write("Inserted revision. ");

        // Delete the first run to create a deletion revision.
        // The first run contains "Original text. ".
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // At this point we have two revisions:
        // 0 – Deletion of the original run, 1 – Insertion of the new text.
        if (doc.Revisions.Count < 2)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Get the deletion revision (index 0) and reject it.
        Revision deletionRevision = doc.Revisions[0];
        deletionRevision.Reject();

        // Attempt to accept the same revision again – it has already been rejected.
        try
        {
            // This will throw because the revision no longer exists in the collection.
            deletionRevision.Accept();
        }
        catch (Exception ex)
        {
            // Handle the error gracefully.
            Console.WriteLine("Error while accepting a rejected revision: " + ex.Message);
        }

        // Accept the remaining insertion revision to finalize the document.
        doc.Revisions[0].Accept();

        // Save the document to verify the final state.
        string outputPath = System.IO.Path.Combine(Environment.CurrentDirectory, "TrackedChangesResult.docx");
        doc.Save(outputPath);
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
