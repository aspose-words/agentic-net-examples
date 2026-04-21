using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start tracking revisions with a specified author and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a paragraph while tracking is enabled – this will be recorded as a revision.
        builder.Writeln("This paragraph is inserted while tracking revisions.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Save the document to verify that the revision was recorded.
        doc.Save("TrackedDocument.docx");

        // Optional: output the number of revisions created.
        Console.WriteLine($"Revisions count: {doc.Revisions.Count}");
    }
}
