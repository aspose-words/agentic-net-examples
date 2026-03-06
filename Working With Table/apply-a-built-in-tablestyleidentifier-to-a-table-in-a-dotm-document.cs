using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template.
        Document doc = new Document("Template.dotm");

        // Create a DocumentBuilder to edit the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell (required before any table formatting).
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Insert a second row with sample data.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.EndRow();

        // Apply a built‑in table style using the StyleIdentifier property.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands;

        // Adjust the table width to fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Finish the table.
        builder.EndTable();

        // Save the modified document as a DOTM file.
        doc.Save("Result.dotm");
    }
}
