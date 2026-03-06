using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the DOTM template document.
        Document doc = new Document("input.dotm");

        // Get all tables in the document (including nested ones).
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any).
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document after inspection (no modifications made).
        doc.Save("output.docx");
    }
}
