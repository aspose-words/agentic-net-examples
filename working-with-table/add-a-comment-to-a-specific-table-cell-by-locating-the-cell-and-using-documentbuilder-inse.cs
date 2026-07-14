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

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable();

        // Locate the target cell (first row, second column).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell targetCell = table.Rows[0].Cells[1];

        // Create a comment.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        // Add a paragraph to the comment to hold the comment text.
        comment.AppendChild(new Paragraph(doc));

        // Create comment range start and end.
        CommentRangeStart rangeStart = new CommentRangeStart(doc, comment.Id);
        CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, comment.Id);

        // Insert the comment range around the cell's first paragraph.
        Paragraph firstParagraph = targetCell.FirstParagraph;
        firstParagraph.PrependChild(rangeStart);
        firstParagraph.AppendChild(rangeEnd);
        // Append the comment node after the range end.
        firstParagraph.AppendChild(comment);

        // Write the comment text.
        builder.MoveTo(comment.FirstParagraph);
        builder.Write("This is a comment on the cell.");

        // Save the document.
        string outputPath = "CommentInTableCell.docx";
        doc.Save(outputPath);
    }
}
