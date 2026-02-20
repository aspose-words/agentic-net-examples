using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load a Markdown document that contains tables.
        // Replace the path with the actual location of your .md file.
        Document doc = new Document(@"Input.md");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if the table has at least one row).
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
