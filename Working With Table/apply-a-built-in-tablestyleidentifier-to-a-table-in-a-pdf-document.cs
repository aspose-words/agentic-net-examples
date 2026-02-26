using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first row (required before applying any table formatting).
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Insert a second row with sample data.
        builder.InsertCell();
        builder.Writeln("Data");
        builder.EndRow();

        // Apply a built‑in table style using the StyleIdentifier property.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Adjust column widths to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Finish the table.
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("TableWithStyle.pdf", SaveFormat.Pdf);
    }
}
