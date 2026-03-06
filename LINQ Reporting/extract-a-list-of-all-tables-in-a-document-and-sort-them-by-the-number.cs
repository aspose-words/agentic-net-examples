using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX document.
        // Replace the path with the actual location of your template.
        Document doc = new Document("InputTemplate.dotx");

        // Retrieve all tables in the document (including nested tables).
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

        // Convert the NodeCollection to a strongly‑typed list of Table objects.
        List<Table> tables = new List<Table>();
        foreach (Table tbl in tableNodes.OfType<Table>())
            tables.Add(tbl);

        // Sort the tables by the number of rows they contain (ascending order).
        List<Table> sortedTables = tables
            .OrderBy(t => t.Rows.Count)   // Change to OrderByDescending for descending order.
            .ToList();

        // Output the sorted list: table index in the original collection and its row count.
        Console.WriteLine("Tables sorted by row count (ascending):");
        for (int i = 0; i < sortedTables.Count; i++)
        {
            Table tbl = sortedTables[i];
            // Find the original index of this table for reference (optional).
            int originalIndex = tables.IndexOf(tbl);
            Console.WriteLine($"SortedIndex: {i}, OriginalIndex: {originalIndex}, RowCount: {tbl.Rows.Count}");
        }

        // (Optional) Save the document if any modifications were made.
        // doc.Save("OutputDocument.dotx");
    }
}
