using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1‑row, 1‑column table (a table must have at least one row before formatting).
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Sample cell");
        builder.EndRow();
        builder.EndTable();

        // Optionally assign a built‑in style to the table.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Apply the desired TableStyleOptions flags (e.g., first column and row banding).
        table.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.RowBands;

        // Save the document as a plain‑text file. The table will be rendered as text.
        doc.Save("TableWithStyle.txt", SaveFormat.Text);
    }
}
