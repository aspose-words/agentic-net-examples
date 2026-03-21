using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddCommentToTableCell
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 2 rows and 3 columns.
        builder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Move the builder's cursor to the end of the target cell:
        // first table (0), second row (1), third column (2).
        builder.MoveToCell(0, 1, 2, -1);

        // Ensure the cell contains at least one paragraph.
        if (builder.CurrentParagraph == null)
        {
            builder.Writeln();
        }

        // Create a new comment anchored to the current paragraph.
        Comment comment = new Comment(doc, "Reviewer", "RV", DateTime.Now);
        comment.SetText("Please verify the data in this cell.");

        // Append the comment to the paragraph where the builder is positioned.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
