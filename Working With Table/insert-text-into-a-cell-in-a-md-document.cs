using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 2x2 table.
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

        // Move the cursor to the cell at row 1, column 1 (second row, second column).
        // Table index is 0 because we have only one table in the document.
        // characterIndex = -1 moves to the end of the cell.
        builder.MoveToCell(0, 1, 1, -1);

        // Insert additional text into the selected cell.
        builder.Write(" – Inserted Text");

        // Save the document (you can change the format by using a different file extension).
        doc.Save("Output.docx");
    }
}
