using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDot
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to simplify table creation.
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
        table.Description = "A simple 2x2 table inserted into a DOT template.";

        // Save the document as a DOT (Word template) file.
        doc.Save("TableTemplate.dot");
    }
}
