using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert at least one cell so that the table can accept formatting.
        builder.InsertCell();
        builder.Write("Header");

        // End the first row.
        builder.EndRow();

        // Add a second row with some data.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.EndRow();

        // Set a built‑in table style (any style that supports conditional formatting).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific style options to the table.
        // For example, enable first row formatting and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Optionally adjust the table to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // End the table construction.
        builder.EndTable();

        // Save the document as an RTF file.
        doc.Save("TableWithStyleOptions.rtf");
    }
}
