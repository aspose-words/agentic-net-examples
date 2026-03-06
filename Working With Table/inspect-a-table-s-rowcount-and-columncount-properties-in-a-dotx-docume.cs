using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("Template.dotx");

        // Retrieve all tables in the document (including nested tables).
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is derived from the number of cells in the first row,
            // assuming the table is well‑formed (all rows have the same number of cells).
            int columnCount = 0;
            if (rowCount > 0)
                columnCount = table.Rows[0].Cells.Count;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
