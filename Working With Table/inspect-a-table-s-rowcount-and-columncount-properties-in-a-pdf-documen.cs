using System;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Load the PDF document.
        Document doc = new Document("input.pdf");

        // Retrieve all tables in the document (including those in all sections).
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
        List<Table> tables = new List<Table>();
        foreach (Node node in tableNodes)
        {
            if (node is Table tbl)
                tables.Add(tbl);
        }

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any rows exist).
            int columnCount = rowCount > 0 ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
