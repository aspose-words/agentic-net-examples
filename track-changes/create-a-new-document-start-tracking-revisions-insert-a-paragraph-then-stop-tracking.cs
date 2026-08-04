using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable tracking of revisions.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert a paragraph while tracking is active.
        builder.Writeln("This paragraph is inserted as a revision.");

        // Disable further tracking.
        doc.StopTrackRevisions();

        // Ensure that a revision was actually recorded.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were recorded.");

        // Save the document.
        doc.Save("TrackedRevisions.docx");
    }
}
