using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOC document. Adjust the path as necessary.
        Document doc = new Document("Tables.docx");

        // Retrieve all tables from the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is typically the number of cells in the first row.
            int columnCount = 0;
            if (rowCount > 0)
                columnCount = table.Rows[0].Cells.Count;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
