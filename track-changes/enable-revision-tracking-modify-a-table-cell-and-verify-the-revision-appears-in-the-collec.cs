using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class RevisionTrackingExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2‑cell table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Original Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Original Cell 2");
        builder.EndTable();

        // Enable revision tracking.
        string author = "Demo Author";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Modify the text of the first cell.
        // Deleting the existing run creates a deletion revision.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell firstCell = table.Rows[0].Cells[0];
        Run existingRun = firstCell.FirstParagraph.Runs[0];
        existingRun.Remove();

        // Inserting a new run creates an insertion revision.
        Run newRun = new Run(doc, "Modified Cell 1");
        firstCell.FirstParagraph.Runs.Add(newRun);

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that revisions were recorded.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after modifying the table cell.");
        }

        // Output basic information about the first revision.
        Revision firstRevision = doc.Revisions[0];
        Console.WriteLine($"Revision detected: Type={firstRevision.RevisionType}, Author={firstRevision.Author}, Date={firstRevision.DateTime}");

        // Save the document to disk.
        doc.Save("RevisionsTable.docx");
    }
}
