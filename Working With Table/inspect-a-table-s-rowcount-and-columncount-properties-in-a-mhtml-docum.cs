using Aspose.Words;
using Aspose.Words.Tables;
using System;

class Program
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("input.mht");

        // Access the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns is the number of cells in the first row (if the table has rows).
            int columnCount = table.FirstRow?.Cells.Count ?? 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
