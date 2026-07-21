using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content that is not a revision.
        builder.Write("This is the original content. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Make changes that will be recorded as revisions.
        builder.Write("This text is added while tracking. ");

        // Stop tracking to finish creating revisions.
        doc.StopTrackRevisions();

        // At this point the document should contain revisions.
        if (!doc.HasRevisions)
            throw new Exception("Expected revisions were not created.");

        // Reject all revisions in the document.
        doc.Revisions.RejectAll();

        // Verify that no revisions remain.
        if (doc.HasRevisions)
            throw new Exception("Revisions were not fully rejected.");

        // Save the resulting document (optional, demonstrates file output).
        doc.Save("RejectedRevisions.docx");

        // Indicate success.
        Console.WriteLine("All tracked changes were rejected; HasRevisions = false.");
    }
}
