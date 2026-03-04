using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTM document.
        Document doc = new Document("Input.dotm");

        // Get all Table nodes in the document (including nested tables).
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
        List<Table> tables = new List<Table>();
        foreach (Table tbl in tableNodes)
            tables.Add(tbl);

        // Sort tables by the number of rows they contain (ascending order).
        List<Table> sortedTables = tables.OrderBy(t => t.Rows.Count).ToList();

        // Output the sorted list: index in the sorted list and row count.
        for (int i = 0; i < sortedTables.Count; i++)
        {
            Console.WriteLine($"Table {i}: Row count = {sortedTables[i].Rows.Count}");
        }

        // Save the (unchanged) document if needed.
        doc.Save("Output.docx");
    }
}
