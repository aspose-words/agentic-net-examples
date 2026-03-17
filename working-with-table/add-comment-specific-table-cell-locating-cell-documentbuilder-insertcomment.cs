using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddCommentToTableCell
{
    static void Main()
    {
        // Load an existing document that contains a table.
        Document doc = new Document("Input.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the target cell location (zero‑based indexes):
        // first table (0), second row (1), third column (2).
        int tableIndex = 0;
        int rowIndex = 1;
        int columnIndex = 2;

        // Move the builder's cursor to the end of the specified cell.
        // The characterIndex of -1 positions the cursor after the last character in the cell.
        builder.MoveToCell(tableIndex, rowIndex, columnIndex, -1);

        // Ensure the cell contains at least one paragraph (required for a comment anchor).
        // If the cell is empty, Insert a new paragraph.
        if (builder.CurrentParagraph == null)
        {
            builder.Writeln();
        }

        // Create a new comment anchored to the current paragraph.
        Comment comment = new Comment(doc, "Reviewer", "RV", DateTime.Now);
        comment.SetText("Please verify the data in this cell.");

        // Append the comment to the paragraph that the builder is currently positioned in.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
