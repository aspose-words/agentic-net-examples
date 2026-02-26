using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InspectTableDimensions
{
    static void Main()
    {
        // Load the PDF document (Aspose.Words can open PDF as a Word document).
        Document doc = new Document("Input.pdf");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is the number of cells in the first row (if any rows exist).
            int columnCount = rowCount > 0 ? table.FirstRow.Count : 0;

            Console.WriteLine($"Table {i + 1}: {rowCount} rows x {columnCount} columns");
        }

        // Optionally save the document (unchanged) to demonstrate the required save lifecycle.
        doc.Save("Output.pdf");
    }
}
