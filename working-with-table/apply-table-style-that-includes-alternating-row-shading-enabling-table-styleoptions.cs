using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table. Insert a first cell to ensure the table has at least one row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Add some data rows.
        for (int i = 1; i <= 4; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln($"{i * 10}");
            builder.EndRow();
        }

        // Finish table construction.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable row banding (alternating row shading) for the table.
        table.StyleOptions = TableStyleOptions.RowBands;

        // Keep the first row formatting (header) as well.
        table.StyleOptions |= TableStyleOptions.FirstRow;

        // Save the document.
        doc.Save("AlternatingRowShading.docx");
    }
}
