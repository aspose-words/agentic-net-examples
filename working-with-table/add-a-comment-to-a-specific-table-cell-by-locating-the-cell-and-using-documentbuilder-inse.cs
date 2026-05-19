using System;
using System.IO;
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
            builder.Write("Cell 0,0");
            builder.InsertCell();
            builder.Write("Cell 0,1");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1,0");
            builder.InsertCell();
            builder.Write("Cell 1,1");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Move the cursor to the cell at row 1, column 0 (second row, first column).
            builder.MoveToCell(0, 1, 0, -1);

            // Create a comment.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("Comment added to this cell.");

            // Insert the comment range and the comment into the current paragraph.
            Paragraph para = builder.CurrentParagraph;

            // Start of comment range.
            CommentRangeStart rangeStart = new CommentRangeStart(doc, comment.Id);
            para.AppendChild(rangeStart);

            // Text that the comment refers to.
            Run commentedRun = new Run(doc, "Cell 1,0");
            para.AppendChild(commentedRun);

            // End of comment range.
            CommentRangeEnd rangeEnd = new CommentRangeEnd(doc, comment.Id);
            para.AppendChild(rangeEnd);

            // The comment itself.
            para.AppendChild(comment);

            // Define the output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CommentInTableCell.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
