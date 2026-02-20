using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("Input.mht");

        // Locate the first table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Choose the cell you want to edit (e.g., first row, first column).
        Cell targetCell = table.Rows[0].Cells[0];

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the first paragraph of the target cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text into the cell.");

        // Save the modified document back to MHTML format.
        doc.Save("Output.mht");
    }
}
