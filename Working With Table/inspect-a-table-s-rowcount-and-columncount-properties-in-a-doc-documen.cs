using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableInfoExample
{
    static void Main()
    {
        // Load an existing DOC document.
        // Replace "Input.doc" with the path to your document.
        Document doc = new Document("Input.doc");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is typically the number of cells in the first row.
            // Guard against empty tables.
            int columnCount = 0;
            if (rowCount > 0)
                columnCount = table.FirstRow.Cells.Count;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally, save the document after inspection (no changes made here).
        // Replace "Output.doc" with the desired output path.
        doc.Save("Output.doc");
    }
}
