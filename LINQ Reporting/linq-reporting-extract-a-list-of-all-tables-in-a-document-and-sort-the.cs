using System;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Retrieve all tables in the document.
        NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);
        List<Table> tables = tableNodes.Cast<Table>().ToList();

        // Sort tables by the number of rows (ascending).
        List<Table> sortedTables = tables
            .OrderBy(t => t.Rows.Count)
            .ToList();

        // Output the sorted information.
        for (int i = 0; i < sortedTables.Count; i++)
        {
            Table table = sortedTables[i];
            Console.WriteLine($"Table {i + 1}: {table.Rows.Count} rows");
        }

        // Optionally, save the document (unchanged) to a new file.
        doc.Save("Result.docx");
    }
}
