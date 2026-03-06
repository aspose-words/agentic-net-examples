using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The method returns the created Table node.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1,1");                // Add text to the cell.

        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 1,2");
        builder.EndRow();                         // End the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Cell 2,1");

        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Cell 2,2");
        builder.EndRow();                         // End the second row.

        // Finish the table construction.
        builder.EndTable();

        // Optional: set alternative text for accessibility.
        table.Title = "Sample Table";
        table.Description = "A table inserted into a DOT template";

        // Save the document as a Word template (.dot).
        doc.Save("TableTemplate.dot");
    }
}
