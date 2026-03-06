using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the MHTML document.
        Document doc = new Document("input.mhtml");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];
            int rowCount = table.Rows.Count;

            // Column count is based on the number of cells in the first row (if any rows exist).
            int columnCount = rowCount > 0 ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally, save the document after inspection (no modifications made).
        doc.Save("output.mhtml");
    }
}
