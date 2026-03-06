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

        // Start building a table. At least one cell must be inserted before any formatting.
        Table table = builder.StartTable();
        builder.InsertCell(); // First cell of the first row.

        // Optionally set a built‑in style for the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific TableStyleOptions flags (e.g., FirstRow and RowBands).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Add some content to the table so it has a visible structure.
        builder.Writeln("Header");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Item 1");
        builder.InsertCell();
        builder.Writeln("10");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Item 2");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document in DOC format.
        doc.Save("TableStyleOptions.doc");
    }
}
