using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDot
{
    static void Main()
    {
        // Load the DOT template (DOT is a Word template file)
        Document doc = new Document("Template.dot");

        // Create a DocumentBuilder attached to the loaded document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to the location where the table should be inserted.
        // For this example we insert at the end of the first section's body.
        builder.MoveToDocumentEnd();

        // Start a new table
        Table table = builder.StartTable();

        // First row, first cell
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // First row, second cell
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row, first cell
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        // Second row, second cell
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table
        builder.EndTable();

        // Optional: set table title and description (useful for DOCX compliance)
        table.Title = "Sample Table";
        table.Description = "A table inserted into a DOT document.";

        // Save the modified document (can be saved as DOCX, DOC, or another DOT)
        doc.Save("Result.docx");
    }
}
