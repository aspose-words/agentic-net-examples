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

        // Build a simple 1x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Enable revision tracking.
        doc.StartTrackRevisions("DemoAuthor", DateTime.Now);

        // Modify the text of the first cell to generate a revision.
        Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell firstCell = firstTable.Rows[0].Cells[0];
        // Clear existing paragraphs and add new text.
        firstCell.RemoveAllChildren();
        firstCell.AppendChild(new Paragraph(doc));
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Modified Cell 1"));

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document (optional for verification).
        doc.Save("TrackedDocument.docx");

        // Verify that a revision was created.
        RevisionCollection revisions = doc.Revisions;
        if (revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after modifying the table cell.");
        }

        // Output revision details.
        Console.WriteLine($"Total revisions: {revisions.Count}");
        foreach (Revision rev in revisions)
        {
            Console.WriteLine($"Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }
    }
}
