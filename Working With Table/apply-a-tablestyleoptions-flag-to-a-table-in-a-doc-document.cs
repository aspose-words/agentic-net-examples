using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table. At least one row must exist before applying formatting.
        Table table = builder.StartTable();
        builder.InsertCell(); // first cell

        // Optionally set a built‑in style identifier for the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply the desired TableStyleOptions flags.
        // Here we enable first column formatting, row banding, and first row formatting.
        table.StyleOptions = TableStyleOptions.FirstColumn |
                              TableStyleOptions.RowBands |
                              TableStyleOptions.FirstRow;

        // Auto‑fit the table to its contents (optional but common).
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Populate the table with sample data.
        builder.Writeln("Item");
        builder.CellFormat.RightPadding = 40;
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

        builder.InsertCell();
        builder.Writeln("Carrots");
        builder.InsertCell();
        builder.Writeln("50");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("TableWithStyleOptions.doc");
    }
}
