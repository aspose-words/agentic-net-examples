using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class ApplyTableStyleOptionsToPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert at least one cell/row before applying any formatting.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add a few data rows.
        builder.InsertCell();
        builder.Write("Row 1, Col 1");
        builder.InsertCell();
        builder.Write("Row 1, Col 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Row 2, Col 1");
        builder.InsertCell();
        builder.Write("Row 2, Col 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific style options (e.g., first row, first column, and row banding).
        table.StyleOptions = TableStyleOptions.FirstRow |
                              TableStyleOptions.FirstColumn |
                              TableStyleOptions.RowBands;

        // Optional: Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a PDF file.
        doc.Save("TableWithStyleOptions.pdf");
    }
}
