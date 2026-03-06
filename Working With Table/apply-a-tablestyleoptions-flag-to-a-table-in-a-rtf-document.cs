using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table and insert the first cell (required before any formatting).
        Table table = builder.StartTable();
        builder.InsertCell();

        // End the first (header) row.
        builder.EndRow();

        // Apply a built‑in table style (optional, but demonstrates style usage).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply the desired TableStyleOptions flags to the table.
        // Example: apply formatting to the first row and enable row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Add a second row with two cells as sample content.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        // Add a third row with sample data.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as RTF.
        doc.Save("TableWithStyle.rtf");
    }
}
