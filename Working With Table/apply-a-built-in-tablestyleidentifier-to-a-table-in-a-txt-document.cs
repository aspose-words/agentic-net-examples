using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToTxt
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first cell (required before setting any table formatting).
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        // Add a second row.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Add a third row.
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("40");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style using the StyleIdentifier property.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally apply style options (first row, first column, row bands).
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a plain‑text file. Table formatting will be lost in the TXT output,
        // but the style is applied in the document model before saving.
        doc.Save("TableWithStyle.txt");
    }
}
