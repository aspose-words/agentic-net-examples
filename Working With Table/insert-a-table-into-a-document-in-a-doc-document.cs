using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow(); // End of first row.

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Cell 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow(); // End of second row.

        // Finish the table.
        builder.EndTable();

        // Optional: Apply a simple style and auto‑fit.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document in DOC format.
        doc.Save("InsertedTable.doc");
    }
}
