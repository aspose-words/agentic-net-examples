using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableInspector
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Get all tables in the document (including nested tables).
        NodeCollection tables = doc.GetChildNodes(NodeType.Table, true);

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = (Table)tables[i];

            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns is taken from the first row's cell count.
            // If the table has no rows, column count is zero.
            int columnCount = rowCount > 0 ? table.Rows[0].Cells.Count : 0;

            Console.WriteLine($"Table #{i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (no modifications made, but required by lifecycle rules).
        doc.Save("Output.docx");
    }
}
