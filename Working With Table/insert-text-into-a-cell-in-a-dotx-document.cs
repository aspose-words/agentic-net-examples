using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Locate the first table in the document (adjust indices as needed).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("No table found in the document.");

        // Select the target cell (e.g., first row, first column).
        Cell targetCell = table.Rows[0].Cells[0];

        // Move the builder's cursor to the first paragraph of the target cell.
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text into the cell.
        builder.Write("Inserted text goes here.");

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
