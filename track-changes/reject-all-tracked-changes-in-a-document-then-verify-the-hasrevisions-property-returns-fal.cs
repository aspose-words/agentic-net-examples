using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content.
        builder.Writeln("Original paragraph.");

        // Enable track changes.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Make some modifications that will be tracked as revisions.
        builder.Writeln("This paragraph was added while tracking changes.");
        builder.MoveTo(doc.FirstSection.Body.FirstParagraph);
        builder.Write("Modified ");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // At this point the document should contain revisions.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Reject all tracked changes using the RevisionCollection API.
        doc.Revisions.RejectAll();

        // Verify that no revisions remain.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully rejected.");

        // Save the final document (optional, demonstrates that the file is valid).
        doc.Save("RejectedDocument.docx");
    }
}
