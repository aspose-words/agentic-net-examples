using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("Template.dotx");

        // Get all tables in the document. GetChildNodes returns a NodeCollection, which does not have ToArray().
        // Use LINQ's Cast<T>() (or OfType<T>()) to enumerate the Table nodes.
        var tables = doc.GetChildNodes(NodeType.Table, true).Cast<Table>();

        // Iterate through each table and output its row and column counts.
        int i = 0;
        foreach (Table table in tables)
        {
            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any rows exist).
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
            i++;
        }

        // Optionally save the document after inspection (no changes made).
        doc.Save("InspectResult.docx");
    }
}
