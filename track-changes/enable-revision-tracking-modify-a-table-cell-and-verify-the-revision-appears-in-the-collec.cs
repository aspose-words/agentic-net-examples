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

        // Insert a table with one row and two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original"); // First cell content.
        builder.InsertCell();
        builder.Write("Second Cell");
        builder.EndRow();
        builder.EndTable();

        // Enable revision tracking.
        doc.StartTrackRevisions("Demo Author", DateTime.Now);

        // Modify the text in the first cell.
        Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell firstCell = firstTable.FirstRow.FirstCell;
        // Replace the existing text with new text.
        if (firstCell.FirstParagraph.Runs.Count > 0)
        {
            firstCell.FirstParagraph.Runs[0].Text = "Modified";
        }
        else
        {
            // If there are no runs, simply write new text.
            firstCell.FirstParagraph.AppendChild(new Run(doc, "Modified"));
        }

        // Stop revision tracking.
        doc.StopTrackRevisions();

        // Verify that at least one revision was created.
        RevisionCollection revisions = doc.Revisions;
        if (revisions.Count == 0)
        {
            throw new Exception("Expected revisions were not found.");
        }

        // Optional: output revision details to the console.
        foreach (Revision rev in revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
        }

        // Save the document to verify the changes and revisions are persisted.
        doc.Save("TrackedDocument.docx");
    }
}
