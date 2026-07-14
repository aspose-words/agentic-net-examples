using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking of revisions with an author name and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Insert a paragraph while tracking is active – this will be recorded as a revision.
        builder.Writeln("This paragraph is inserted while tracking changes.");

        // Disable further tracking of revisions.
        doc.StopTrackRevisions();

        // Ensure that at least one revision was captured.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were recorded.");
        }

        // Save the resulting document to a file.
        doc.Save("TrackChanges.docx");
    }
}
