using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDocm
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2, Row 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Optionally set table title and description (useful for accessibility).
        table.Title = "Sample Table";
        table.Description = "A simple 2x2 table inserted into a DOCM document.";

        // Save the document as a macro-enabled DOCM file.
        doc.Save("TableInDocument.docm", SaveFormat.Docm);
    }
}
