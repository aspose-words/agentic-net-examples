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
        builder.Writeln("Original text. ");

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert new text – this will be recorded as an insertion revision.
        builder.Writeln("Inserted revision text. ");

        // Delete the first run (the original text) – this will be recorded as a deletion revision.
        doc.FirstSection.Body.FirstParagraph.Runs[0].Remove();

        // Stop tracking further changes (optional).
        doc.StopTrackRevisions();

        // At this point the document should have revisions.
        if (!doc.HasRevisions)
            throw new Exception("Expected revisions were not created.");

        // Reject all revisions in the document.
        doc.Revisions.RejectAll();

        // Verify that no revisions remain.
        if (doc.HasRevisions)
            throw new Exception("Revisions were not successfully rejected.");

        // Save the resulting document.
        doc.Save("RejectedRevisions.docx");
    }
}
