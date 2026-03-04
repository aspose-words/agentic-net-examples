using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableInspector
{
    static void Main()
    {
        // Load the HTML document (replace with your actual file path)
        Document doc = new Document("input.html");

        // Get the collection of tables in the first section's body
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any rows exist)
            int columnCount = rowCount > 0 ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i}: RowCount = {rowCount}, ColumnCount = {columnCount}");
        }

        // Optionally save the document after inspection (replace with desired output path)
        doc.Save("output.html");
    }
}
