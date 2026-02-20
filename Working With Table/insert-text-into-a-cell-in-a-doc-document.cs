using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();               // Row 1, Cell 1
        builder.Write("R1C1");
        builder.InsertCell();               // Row 1, Cell 2
        builder.Write("R1C2");
        builder.EndRow();

        builder.InsertCell();               // Row 2, Cell 1
        builder.Write("R2C1");
        builder.InsertCell();               // Row 2, Cell 2
        builder.Write("R2C2");
        builder.EndRow();
        builder.EndTable();

        // Choose the cell we want to modify (e.g., second row, first column).
        Cell targetCell = table.Rows[1].Cells[0]; // Zero‑based indexing

        // Move the builder's cursor to the first paragraph of the target cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text");

        // Save the document to a file.
        doc.Save("Result.docx");
    }
}
