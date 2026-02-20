using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ApplyTableStyleToPdf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first row (required before setting any table formatting).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Insert a second row.
        builder.InsertCell();
        builder.Writeln("Value 1");
        builder.InsertCell();
        builder.Writeln("Value 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Apply a built‑in table style by its identifier.
        // For example, use the LightGrid style.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally, specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Save the document as PDF.
        string outputPath = "TableWithStyle.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);
    }
}
