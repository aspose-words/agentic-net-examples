using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToRtf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell – a table must contain at least one row before styling.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        // Add a second row with sample data.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Apply a built‑in table style using the StyleIdentifier property.
        // Any style from the StyleIdentifier enum can be used; here we choose LightGrid.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.RowBands |
                              TableStyleOptions.FirstColumn;

        // Auto‑fit the table to its contents so the style is rendered correctly.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // End the table construction.
        builder.EndTable();

        // Save the document as an RTF file.
        doc.Save("TableWithStyle.rtf");
    }
}
