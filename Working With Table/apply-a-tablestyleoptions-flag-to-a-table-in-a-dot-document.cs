using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document (DOT template)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and insert at least one cell (required before setting formatting)
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Apply a built‑in table style
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply specific style options (e.g., first row and row banding)
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // End the table
        builder.EndTable();

        // Save the document as a DOT template
        doc.Save("TableWithStyleOptions.dot");
    }
}
