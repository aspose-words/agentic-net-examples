using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOT template.
        Document doc = new Document("Template.dot");

        // Create a DocumentBuilder attached to the loaded document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Position the builder at the start of the main story (the document body).
        builder.MoveToDocumentStart();

        // Begin a new table.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1, Row 1");           // Add text to the cell.

        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 2, Row 1");

        builder.EndRow();                         // End the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Cell 1, Row 2");

        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Cell 2, Row 2");

        builder.EndRow();                         // End the second row.

        // Finish the table.
        builder.EndTable();

        // Optional: set title and description for accessibility.
        table.Title = "Sample Table";
        table.Description = "Demonstrates inserting a table into a DOT document.";

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
