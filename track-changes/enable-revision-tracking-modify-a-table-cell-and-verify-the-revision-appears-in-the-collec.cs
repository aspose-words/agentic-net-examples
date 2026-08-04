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

        // Insert a simple 2‑cell table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Original Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Original Cell 2");
        builder.EndTable();

        // Enable revision tracking.
        string author = "Test Author";
        DateTime revisionDate = DateTime.Now;
        doc.StartTrackRevisions(author, revisionDate);

        // Modify the text of the first cell to generate a revision.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell firstCell = table.Rows[0].Cells[0];
        // Clear existing runs.
        firstCell.FirstParagraph.Runs.Clear();
        // Write new text – this will be recorded as an insertion revision.
        builder.MoveTo(firstCell.FirstParagraph);
        builder.Write("Modified Cell 1");

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that at least one revision exists and that it is an insertion.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were created.");

        bool insertionFound = false;
        foreach (Revision rev in doc.Revisions)
        {
            if (rev.RevisionType == RevisionType.Insertion && rev.Author == author)
            {
                insertionFound = true;
                break;
            }
        }

        if (!insertionFound)
            throw new InvalidOperationException("Expected insertion revision not found.");

        // Save the document to verify the result.
        doc.Save("RevisionsExample.docx");
    }
}
