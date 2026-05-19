using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that will NOT be a revision.
        builder.Writeln("Original content. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Add content that will be recorded as a revision.
        builder.Writeln("This is a tracked change. ");

        // Stop tracking so further edits are not recorded.
        doc.StopTrackRevisions();

        // At this point the document should have revisions.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Reject all revisions, reverting the document to its original state.
        doc.Revisions.RejectAll();

        // Verify that no revisions remain.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully rejected.");

        // Save the resulting document.
        doc.Save("RejectedRevisions.docx");

        // Indicate success.
        Console.WriteLine("All revisions rejected successfully; HasRevisions = " + doc.HasRevisions);
    }
}
