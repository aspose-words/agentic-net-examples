using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoDot
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
        builder.Write("Header 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Data 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Data 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Optionally apply a built‑in style to the table.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a DOT (Word template) file.
        doc.Save("Table.dot");
    }
}
