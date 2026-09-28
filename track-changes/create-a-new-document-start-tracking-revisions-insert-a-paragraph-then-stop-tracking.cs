using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Author", DateTime.Now);

        // Insert a paragraph while tracking is active.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph is inserted with tracking enabled.");

        // Disable tracking of revisions.
        doc.StopTrackRevisions();

        // Save the document to a file.
        doc.Save("TrackChanges.docx");
    }
}
