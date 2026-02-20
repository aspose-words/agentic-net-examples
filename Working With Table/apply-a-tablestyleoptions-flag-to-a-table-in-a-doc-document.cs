using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first cell (required before setting any table formatting).
        builder.InsertCell();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific style options using the TableStyleOptions flags.
        // Here we enable formatting for the first row, the last row, and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.LastRow |
                             TableStyleOptions.RowBands;

        // Optional: let the table auto‑fit to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Populate the first row.
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Populate a data row.
        builder.InsertCell();
        builder.Writeln("Item A");
        builder.InsertCell();
        builder.Writeln("Value A");
        builder.EndRow();

        // Populate another data row.
        builder.InsertCell();
        builder.Writeln("Item B");
        builder.InsertCell();
        builder.Writeln("Value B");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("TableStyleOptions.doc");
    }
}
