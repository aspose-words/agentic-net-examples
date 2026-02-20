using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToTxt
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first row with two cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Insert a second row with two cells.
        builder.InsertCell();
        builder.Writeln("Value 1");
        builder.InsertCell();
        builder.Writeln("Value 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style using its StyleIdentifier.
        // For example, use the "Light Grid" style.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally, enable additional style options (first row, banding, etc.).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document as a plain‑text file.
        // The table will be rendered as plain text, but the style is applied in the document model.
        doc.Save("TableWithStyle.txt");
    }
}
