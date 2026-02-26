using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table. At least one cell must be inserted before any table formatting.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // Apply a built‑in table style by its identifier.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Optionally specify which parts of the style are applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // End the table.
        builder.EndTable();

        // Save the document as plain text. The table will be rendered as tab‑delimited text.
        doc.Save("TableWithStyle.txt");
    }
}
