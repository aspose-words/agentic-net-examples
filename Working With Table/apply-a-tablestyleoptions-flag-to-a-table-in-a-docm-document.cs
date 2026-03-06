using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Insert at least one cell and some content – a table must have a row before formatting.
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow();

        // Optionally assign a built‑in style to the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply the desired TableStyleOptions flags.
        // Here we combine FirstColumn, RowBands and FirstRow as an example.
        table.StyleOptions = TableStyleOptions.FirstColumn |
                             TableStyleOptions.RowBands |
                             TableStyleOptions.FirstRow;

        // Save the document as a DOCM file.
        doc.Save("TableWithStyleOptions.docm");
    }
}
