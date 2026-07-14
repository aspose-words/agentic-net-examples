using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text that is not a revision.
        builder.Writeln("Original content.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Add text while tracking – this will create a revision.
        builder.Writeln("This is a tracked change.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point the document should have revisions.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Reject all revisions in the document.
        doc.Revisions.RejectAll();

        // Verify that there are no remaining revisions.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully rejected.");

        // Save the resulting document (optional, demonstrates saving).
        doc.Save("Result.docx");
    }
}
