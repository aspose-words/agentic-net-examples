using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleOptions
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell (required before any formatting).
        builder.InsertCell();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply desired style options (first column, first row, row banding).
        table.StyleOptions = TableStyleOptions.FirstColumn |
                             TableStyleOptions.FirstRow |
                             TableStyleOptions.RowBands;

        // Auto‑fit the table to its contents.
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

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "ApplyTableStyleOptions.docx";
        doc.Save(outputPath);
    }
}
