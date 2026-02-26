using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInspection
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the RTF document.
            // Replace "input.rtf" with the path to your RTF file.
            Document doc = new Document("input.rtf");

            // Get the collection of tables in the first section's body.
            TableCollection tables = doc.FirstSection.Body.Tables;

            // Iterate through each table and output its row and column counts.
            for (int i = 0; i < tables.Count; i++)
            {
                Table table = tables[i];

                // Row count is the number of rows in the table.
                int rowCount = table.Rows.Count;

                // Column count is the number of cells in the first row (if any rows exist).
                int columnCount = 0;
                if (rowCount > 0)
                {
                    // The first row's cell count represents the number of columns.
                    columnCount = table.FirstRow.Cells.Count;
                }

                Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
            }

            // Optionally, save the document after inspection (no changes made).
            // Replace "output.rtf" with the desired output path.
            doc.Save("output.rtf");
        }
    }
}
