using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder for editing the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking of revisions with an author name and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a paragraph while tracking is active.
        builder.Writeln("This paragraph was added while tracking revisions.");

        // Disable tracking of revisions.
        doc.StopTrackRevisions();

        // Verify that at least one revision was generated.
        if (doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException($"Expected at least 1 revision, but found none.");
        }

        // Save the document to a file.
        doc.Save("TrackedDocument.docx");
    }
}
