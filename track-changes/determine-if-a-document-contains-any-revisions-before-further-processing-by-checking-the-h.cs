using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text without tracking – this should not create revisions.
        builder.Write("Initial content. ");

        // Verify that there are no revisions at this point.
        bool hasRevisions = doc.HasRevisions;
        Console.WriteLine($"Has revisions before tracking? {hasRevisions}");

        // Start tracking revisions with a specific author and timestamp.
        doc.StartTrackRevisions("Jane Doe", DateTime.Now);

        // Any changes made after this call are recorded as revisions.
        builder.Write("Added revision text. ");

        // Stop tracking to avoid further changes being recorded.
        doc.StopTrackRevisions();

        // Check again for revisions.
        hasRevisions = doc.HasRevisions;
        Console.WriteLine($"Has revisions after tracking? {hasRevisions}");

        // Optionally, save the document to inspect the revisions manually.
        doc.Save("RevisionsDemo.docx");
    }
}
