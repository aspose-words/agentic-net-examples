using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
        Table table = builder.StartTable();
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the entire table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;

        // Enable all style options so the style affects the whole table.
        table.StyleOptions = TableStyleOptions.FirstRow |
                             TableStyleOptions.LastRow |
                             TableStyleOptions.FirstColumn |
                             TableStyleOptions.LastColumn |
                             TableStyleOptions.RowBands |
                             TableStyleOptions.ColumnBands;

        // Save the document.
        doc.Save("StyledTable.docx");
    }
}
