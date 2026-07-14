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

        // Insert a 1x1 table and write initial text into the single cell.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original cell text.");
        builder.EndRow();
        builder.EndTable();

        // Enable revision tracking with a specific author and timestamp.
        doc.StartTrackRevisions("John Doe", DateTime.Now);

        // Modify the text inside the existing table cell.
        // Retrieve the first cell of the first row.
        Cell cell = table.FirstRow.FirstCell;
        // Clear existing contents and write new text – this creates revisions.
        cell.RemoveAllChildren();
        builder.MoveTo(cell.FirstParagraph);
        builder.Write("Modified cell text.");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that at least one revision was recorded.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after modifying the table cell.");
        }

        // Optionally, inspect the first revision (should be a deletion of the original text).
        Revision firstRevision = doc.Revisions[0];
        Console.WriteLine($"Revision author: {firstRevision.Author}");
        Console.WriteLine($"Revision type: {firstRevision.RevisionType}");
        Console.WriteLine($"Revision text: {firstRevision.ParentNode.GetText().Trim()}");

        // Save the document to verify the revisions persist.
        doc.Save("TrackedRevisions.docx");
    }
}
