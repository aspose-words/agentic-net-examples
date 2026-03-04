using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Load a TXT document. TxtLoadOptions can be customized if needed.
        Document doc = new Document("input.txt", new TxtLoadOptions());

        // Retrieve all tables in the document.
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and output its row and column counts.
        foreach (Table table in tables)
        {
            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns is determined by the cell count of the first row.
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table found: Rows = {rowCount}, Columns = {columnCount}");
        }
    }
}
