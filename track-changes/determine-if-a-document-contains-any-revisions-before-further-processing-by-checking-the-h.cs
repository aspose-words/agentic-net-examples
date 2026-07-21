using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial content. This does NOT count as a revision.
        builder.Writeln("Original text.");

        // Verify that the document has no revisions at this point.
        bool hasRevisionsBefore = doc.HasRevisions;
        Console.WriteLine($"Has revisions before tracking: {hasRevisionsBefore}");

        // Enable tracking of changes.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Add new content while tracking is enabled – this will be recorded as a revision.
        builder.Writeln("Added revision text.");

        // Stop tracking further changes (optional, but demonstrates lifecycle usage).
        doc.StopTrackRevisions();

        // Verify that the document now reports having revisions.
        bool hasRevisionsAfter = doc.HasRevisions;
        Console.WriteLine($"Has revisions after tracking: {hasRevisionsAfter}");

        // Save the document to disk (demonstrates the required save step).
        doc.Save("RevisionsDemo.docx");
    }
}
