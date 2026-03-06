using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInspectionExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source document (replace with your actual file path).
            Document doc = new Document("InputDocument.docx");

            // Get the collection of tables in the first section's body.
            TableCollection tables = doc.FirstSection.Body.Tables;

            // Iterate through each table and output its row and column counts.
            for (int i = 0; i < tables.Count; i++)
            {
                Table table = tables[i];

                // Row count is the number of Row objects in the table.
                int rowCount = table.Rows.Count;

                // Column count is the number of cells in the first row (if the table has at least one row).
                int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

                Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
            }

            // Optionally, save the document if any modifications were made.
            doc.Save("OutputDocument.docx");
        }
    }
}
