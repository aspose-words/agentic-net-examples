using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the DOT template.
        Document doc = new Document("Template.dot");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if the table has at least one row).
            int columnCount = table.Rows.Count > 0 ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document after inspection.
        doc.Save("InspectResult.docx");
    }
}
