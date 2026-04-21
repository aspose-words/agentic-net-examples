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

        // Insert a simple 1x1 table with initial text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Original");
        builder.EndRow();
        builder.EndTable();

        // Enable revision tracking with a specific author.
        string author = "Demo Author";
        doc.StartTrackRevisions(author, DateTime.Now);

        // Modify the content of the first cell.
        // Clear existing runs and add a new run with modified text.
        Cell firstCell = table.Rows[0].Cells[0];
        firstCell.FirstParagraph.Runs.Clear();
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Modified"));

        // Stop tracking revisions.
        doc.StopTrackRevisions();

        // Save the document so the revisions are persisted.
        string outputPath = "RevisionsDemo.docx";
        doc.Save(outputPath);

        // Verify that at least one revision was created.
        if (!doc.HasRevisions || doc.Revisions.Count == 0)
        {
            throw new InvalidOperationException("No revisions were detected after modifying the table cell.");
        }

        // Iterate through revisions and output basic information.
        Console.WriteLine($"Total revisions: {doc.Revisions.Count}");
        foreach (Revision rev in doc.Revisions)
        {
            Console.WriteLine($"Revision - Type: {rev.RevisionType}, Author: {rev.Author}, Date: {rev.DateTime}");
            // Ensure the revision belongs to the modified cell.
            if (rev.ParentNode != null && rev.ParentNode.GetText().Contains("Modified"))
            {
                Console.WriteLine("Found revision related to the modified cell.");
            }
        }
    }
}
