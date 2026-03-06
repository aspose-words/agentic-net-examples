using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();               // Cell (0,0)
        builder.Write("Cell 0,0");
        builder.InsertCell();               // Cell (0,1)
        builder.Write("Cell 0,1");
        builder.EndRow();

        // Second row.
        builder.InsertCell();               // Cell (1,0)
        builder.Write("Cell 1,0");
        builder.InsertCell();               // Cell (1,1)
        builder.Write("Cell 1,1");
        builder.EndTable();

        // Move the cursor to the first cell of the second row (table index 0, row 1, column 0).
        // characterIndex = 0 positions the cursor at the start of the cell.
        builder.MoveToCell(0, 1, 0, 0);

        // Insert the desired text into that cell.
        builder.Write("Inserted Text");

        // Save the document to disk.
        doc.Save("Output.docx");
    }
}
