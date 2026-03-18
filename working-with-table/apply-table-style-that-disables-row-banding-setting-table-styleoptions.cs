using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class DisableRowBandingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert a single cell so that the table has at least one row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Add a second row with some data.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.EndRow();

        // Apply a built‑in table style (any style will do; here we use MediumShading1Accent1).
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // By default the style may enable row banding (RowBands flag). 
        // To disable row banding we clear that flag. Setting the options to None removes all conditional formatting.
        table.StyleOptions = TableStyleOptions.None; // NoRowBanding equivalent

        // Save the document.
        doc.Save("Table_NoRowBanding.docx");
    }
}
