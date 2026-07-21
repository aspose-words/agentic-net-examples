using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableCommentExample
{
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
            Table table = doc.FirstSection.Body.Tables[0];
            Cell targetCell = table.Rows[0].Cells[1];

            // Move the builder's cursor to the beginning of the target cell.
            builder.MoveTo(targetCell.FirstParagraph);

            // Create a comment and attach it to the paragraph inside the cell.
            Comment comment = new Comment(doc, "Author", "AI", DateTime.Now);
            // Append the comment node to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);
            // Move the builder inside the comment to add the comment text.
            builder.MoveTo(comment.AppendChild(new Paragraph(doc)));
            builder.Write("This is a comment added to the cell.");

            // Save the document.
            string outputPath = "CommentInTableCell.docx";
            doc.Save(outputPath);
        }
    }
}
