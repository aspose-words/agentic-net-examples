using Aspose.Words;
using Aspose.Words.Tables;
using System;

class RevisionDemo
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Begin tracking revisions with a specific author.
        doc.StartTrackRevisions("John Doe");

        // Modify the text of the first cell – this change will be recorded as a revision.
        Cell firstCell = table.FirstRow.FirstCell;
        // Clear existing runs and add new text.
        firstCell.FirstParagraph.Runs.Clear();
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Modified Cell 1"));

        // Stop tracking further changes.
        doc.StopTrackRevisions();

        // Verify that a revision was captured.
        if (doc.HasRevisions && doc.Revisions.Count > 0)
        {
            Revision rev = doc.Revisions[0];
            Console.WriteLine($"Revision author: {rev.Author}");
            Console.WriteLine($"Revision type: {rev.RevisionType}");
        }
        else
        {
            Console.WriteLine("No revisions were recorded.");
        }

        // Save the resulting document.
        doc.Save("RevisionTable.docx");
    }
}
