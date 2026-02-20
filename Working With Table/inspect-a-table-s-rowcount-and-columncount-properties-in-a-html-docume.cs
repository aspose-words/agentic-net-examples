using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document.
        Document doc = new Document("input.html");

        // Get the collection of tables in the document's first section body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns is determined by the cell count of the first row.
            // If the table has no rows, column count is zero.
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
