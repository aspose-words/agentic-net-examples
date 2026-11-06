using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Enable revision tracking.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Modify the first cell while tracking is enabled.
        Cell firstCell = table.FirstRow.FirstCell;
        builder.MoveTo(firstCell.FirstParagraph);
        builder.Writeln("Inserted while tracking");

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document.
        doc.Save("TrackedDocument.docx");

        // Inspect the revisions collection.
        RevisionCollection revisions = doc.Revisions;
        if (revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after modifying the table cell.");
        }

        // Output revision details.
        Console.WriteLine($"Revision count: {revisions.Count}");
        foreach (Revision rev in revisions)
        {
            // RevisionDateTime property may not be available in older versions; omit if not present.
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}");
        }
    }
}
