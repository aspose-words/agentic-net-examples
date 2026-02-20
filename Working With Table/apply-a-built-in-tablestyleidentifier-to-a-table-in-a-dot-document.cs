using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToDot
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell – a table must contain at least one row before any style can be set.
        builder.InsertCell();

        // Apply a built‑in table style by its identifier.
        // Example: LightGrid (you can replace with any other StyleIdentifier value).
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Let the table auto‑fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Populate the table with a simple two‑row example.
        builder.Writeln("Header");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Data");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a DOT (Word template) file.
        doc.Save("TableWithStyle.dot");
    }
}
