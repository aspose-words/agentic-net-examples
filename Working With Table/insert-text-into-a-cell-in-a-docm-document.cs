using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTextIntoCell
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Locate the target cell.
        // For example, get the first cell in the first table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        Cell targetCell = table.Rows[0].Cells[0];

        // Move the builder cursor to the beginning of the cell's first paragraph.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(targetCell.FirstParagraph);

        // Insert the desired text.
        builder.Write("Inserted text into the cell.");

        // Save the modified document.
        doc.Save("OutputDocument.docm");
    }
}
