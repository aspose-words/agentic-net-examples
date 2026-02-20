using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table. At least one cell must be inserted before any formatting.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header");

        // Second cell of the first row.
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Add a second row with sample data.
        builder.InsertCell();
        builder.Writeln("Item1");
        builder.InsertCell();
        builder.Writeln("100");
        builder.EndRow();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Apply specific style options (first row formatting and row banding).
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document as plain text, attempting to preserve the table layout.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };
        doc.Save("TableWithStyle.txt", txtOptions);
    }
}
