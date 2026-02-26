using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDotm
{
    static void Main()
    {
        // Load an existing DOTM template.
        Document doc = new Document("Template.dotm");

        // Create a DocumentBuilder associated with the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
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

        // Optionally set a title and description for the table.
        table.Title = "Sample Table";
        table.Description = "A table inserted into a DOTM document.";

        // Save the modified document back to DOTM format.
        doc.Save("Result.dotm");
    }
}
