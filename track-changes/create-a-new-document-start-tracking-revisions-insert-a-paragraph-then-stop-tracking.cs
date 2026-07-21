using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Normal editing before tracking does not create revisions.
        builder.Writeln("Initial content. ");

        // Start tracking revisions with an author name and the current date.
        doc.StartTrackRevisions("Sample Author", DateTime.Now);

        // Insert a paragraph while tracking is enabled – this will be recorded as a revision.
        builder.Writeln("This paragraph is inserted while tracking changes.");

        // Stop tracking revisions – further edits will not be recorded as revisions.
        doc.StopTrackRevisions();

        // Save the document to a file.
        doc.Save("TrackChangesExample.docx");
    }
}
