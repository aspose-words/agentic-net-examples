using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableInfo
{
    static void Main()
    {
        // Load an existing DOC document.
        Document doc = new Document("Input.doc");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is derived from the first row's cell count (assuming a regular table).
            int columnCount = table.Rows.Count > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally, save the document if any modifications were made.
        // doc.Save("Output.doc");
    }
}
