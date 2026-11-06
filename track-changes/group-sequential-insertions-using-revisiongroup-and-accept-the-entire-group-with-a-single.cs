using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add initial content (no revisions).
        builder.Writeln("Original paragraph.");

        // Start tracking revisions.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Insert sequential lines – these will become separate insert revisions.
        builder.Writeln("First inserted line.");
        builder.Writeln("Second inserted line.");
        builder.Writeln("Third inserted line.");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Verify that revisions were created.
        if (doc.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions but none were found.");

        // Accept all revisions in one call (equivalent to accepting a RevisionGroup).
        doc.AcceptAllRevisions();

        // Save the resulting document.
        doc.Save("GroupedInsertions.docx");
    }
}
