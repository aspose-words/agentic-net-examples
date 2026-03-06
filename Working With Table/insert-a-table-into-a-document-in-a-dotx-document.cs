using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDotx
{
    static void Main()
    {
        // Load an existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to simplify inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1,1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 2,1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optional: set table title and description (useful for accessibility).
        table.Title = "Sample Table";
        table.Description = "A simple 2x2 table inserted into a DOTX document.";

        // Save the modified document as a DOCX file.
        doc.Save("Result.docx");
    }
}
