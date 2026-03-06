using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDocm
{
    static void Main()
    {
        // Load an existing DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table at the current cursor position.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Optionally set a title and description for accessibility.
        table.Title = "Sample Table";
        table.Description = "A simple 2x2 table inserted into a DOCM file.";

        // Save the modified document as a DOCM.
        doc.Save("OutputDocument.docm");
    }
}
