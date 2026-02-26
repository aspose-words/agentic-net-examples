using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("InputDocument.mhtml");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        foreach (Table table in tables)
        {
            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if the table has at least one row).
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (optional, maintains lifecycle compliance).
        doc.Save("OutputDocument.mhtml");
    }
}
