using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some initial text. This does NOT count as a revision.
        builder.Writeln("Initial content without tracking.");

        // Verify that the document currently has no revisions.
        bool hasRevisionsBefore = doc.HasRevisions;
        Console.WriteLine($"Has revisions before tracking? {hasRevisionsBefore}");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Jane Doe", DateTime.Now);

        // Add text while tracking is enabled – this will be recorded as a revision.
        builder.Writeln("This text is added as a revision.");

        // Stop tracking to avoid further changes being recorded.
        doc.StopTrackRevisions();

        // Check the HasRevisions property after making tracked changes.
        bool hasRevisionsAfter = doc.HasRevisions;
        Console.WriteLine($"Has revisions after tracking? {hasRevisionsAfter}");

        // Optionally, save the document to verify the revisions visually in Word.
        doc.Save("TrackedRevisions.docx");
    }
}
