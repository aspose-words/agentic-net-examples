using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableDimensionsInspector
{
    static void Main()
    {
        // Load an existing DOC document.
        // Replace the path with the actual location of your document.
        Document doc = new Document(@"C:\Docs\Sample.doc");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table.
        for (int t = 0; t < tables.Count; t++)
        {
            Table table = tables[t];

            // Row count is the number of rows in the table.
            int rowCount = table.Rows.Count;

            // Column count is taken from the first row's cell count.
            // If the table has no rows, column count is zero.
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {t + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
