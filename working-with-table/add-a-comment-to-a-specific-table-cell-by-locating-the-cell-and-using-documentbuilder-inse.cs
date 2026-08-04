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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1,1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 1,2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 2,1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2,2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Locate the specific cell to comment (first row, second column).
        Table table = doc.FirstSection.Body.Tables[0];
        Cell targetCell = table.Rows[0].Cells[1]; // zero‑based indices.

        // Ensure the cell has at least one paragraph.
        targetCell.EnsureMinimum();
        Paragraph para = targetCell.FirstParagraph;

        // Create a comment.
        Comment comment = new Comment(doc, "Author", "AU", DateTime.Now);
        comment.SetText("This is a comment on the cell.");

        // Insert comment range start before the first run.
        CommentRangeStart rangeStart = new CommentRangeStart(doc, comment.Id);
        para.PrependChild(rangeStart);

        // Insert comment range end after the last run.
        CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, comment.Id);
        para.AppendChild(rangeEnd);

        // Append the comment node itself.
        para.AppendChild(comment);

        // Save the document.
        string outputPath = "CommentInTableCell.docx";
        doc.Save(outputPath);
    }
}
