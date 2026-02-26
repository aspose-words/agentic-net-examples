using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableInspector
{
    static void Main()
    {
        // Load an existing Word document (replace with your actual file path)
        Document doc = new Document("InputDocument.docx");

        // Retrieve all tables in the document (deep search)
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and output its row and column counts
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Row count is the number of Row objects in the table
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if the table has at least one row)
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (unchanged) to demonstrate the required save lifecycle step
        doc.Save("OutputDocument.docx");
    }
}
