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
        builder.InsertCell();

        // Optionally assign a built‑in style to the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply the desired style options (first column, first row, and row banding).
        table.StyleOptions = TableStyleOptions.FirstColumn |
                              TableStyleOptions.FirstRow |
                              TableStyleOptions.RowBands;

        // Populate the first row.
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity (kg)");
        builder.EndRow();

        // Populate a second row.
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Populate a third row.
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("40");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a DOT template.
        doc.Save("TableWithStyle.dot");
    }
}
