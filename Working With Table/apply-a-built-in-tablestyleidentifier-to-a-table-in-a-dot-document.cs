using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document (DOT template)
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table
        Table table = builder.StartTable();

        // Insert first row (required before setting any table formatting)
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Insert second row with sample data
        builder.InsertCell();
        builder.Writeln("Data");
        builder.EndRow();

        // Finish the table
        builder.EndTable();

        // Apply a built‑in table style using its identifier
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Optionally specify which parts of the style to apply
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document as a DOT (Word template) file
        doc.Save("StyledTable.dot", SaveFormat.Dot);
    }
}
