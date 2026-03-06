using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDotx
{
    static void Main()
    {
        // Load an existing DOTX template.
        // Replace "Template.dotx" with the path to your DOTX file.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder attached to the loaded document.
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

        // End the table.
        builder.EndTable();

        // Optionally set a title and description for the table (useful for DOCX compliance).
        table.Title = "Sample Table";
        table.Description = "A simple 2x2 table inserted into a DOTX document.";

        // Save the modified document back as a DOTX file.
        // Replace "Result.dotx" with the desired output path.
        doc.Save("Result.dotx");
    }
}
