using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Locate the first cell (row 0, column 0).
        Table table = doc.FirstSection.Body.Tables[0];
        Cell targetCell = table.Rows[0].Cells[0];

        // Move the builder's cursor to the first paragraph of the target cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Create a comment and attach it to the paragraph.
        Comment comment = new Comment(doc, "Author", "A", DateTime.Now);
        // Append the comment node to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);
        // Move the builder inside the comment to add comment text.
        builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
        builder.Write("Comment on the first cell.");

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithComment.docx");
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not saved correctly.");
        }
    }
}
