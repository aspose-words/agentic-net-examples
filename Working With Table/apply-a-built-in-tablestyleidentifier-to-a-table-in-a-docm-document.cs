using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder to work with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. Insert at least one cell before applying any table formatting.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Apply a built‑in table style using its identifier.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands |
                             TableStyleOptions.FirstRow;

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

        // Finish the table.
        builder.EndTable();

        // Save the document as a DOCM file.
        doc.Save("TableWithStyle.docm");
    }
}
