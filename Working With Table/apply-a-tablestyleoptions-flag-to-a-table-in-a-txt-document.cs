using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. At least one cell must be inserted before any formatting.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Add a second row with data.
        builder.InsertCell();
        builder.Writeln("Data");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply desired style options to the table.
        // Here we enable formatting for the first row and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document as a plain‑text file. Tables are converted to text in TXT format.
        doc.Save("TableWithStyle.txt");
    }
}
