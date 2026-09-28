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

        // Start tracking revisions.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a new paragraph while tracking is enabled.
        builder.Writeln("Inserted paragraph while tracking.");

        // Delete the original paragraph to create a deletion revision.
        Paragraph originalParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
        originalParagraph.Remove();

        // Change formatting of the inserted paragraph to create a format revision.
        builder.Font.Bold = true;
        builder.Writeln("Bold text added.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that revisions exist.
        if (!doc.HasRevisions)
            throw new InvalidOperationException("Expected revisions were not created.");

        // Reject all revisions using the RevisionCollection API.
        doc.Revisions.RejectAll();

        // Verify that no revisions remain.
        if (doc.HasRevisions)
            throw new InvalidOperationException("Revisions were not fully rejected.");

        // Save the resulting document.
        doc.Save("Result.docx");

        // Indicate success.
        Console.WriteLine("All revisions rejected successfully; HasRevisions = false.");
    }
}
