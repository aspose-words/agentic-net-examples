using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoMhtmlCell
{
    static void Main()
    {
        // Load the existing MHTML document.
        Document doc = new Document("input.mhtml");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Locate the first table in the document (adjust the index if needed).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Choose the target cell (e.g., first row, first column).
        Cell targetCell = table.Rows[0].Cells[0];

        // Move the builder's cursor to the beginning of the cell's first paragraph.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text goes here.");

        // Save the modified document back to MHTML format.
        doc.Save("output.mhtml");
    }
}
