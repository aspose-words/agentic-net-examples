using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class TableStyleOptionsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert a cell to satisfy the requirement of having at least one row before formatting.
        builder.InsertCell();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific style options (e.g., first row, first column, and row banding).
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.FirstColumn |
                              TableStyleOptions.RowBands;

        // Populate the table with some sample data.
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("TableStyleOptions.pdf", SaveFormat.Pdf);
    }
}
