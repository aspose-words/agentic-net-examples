using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableInspectionExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing DOCM document.
            // Replace "Input.docm" with the path to your DOCM file.
            Document doc = new Document("Input.docm");

            // Iterate through all tables in the document.
            foreach (Table table in doc.GetChildNodes(NodeType.Table, true))
            {
                // Row count is the number of Row objects in the table.
                int rowCount = table.Rows.Count;

                // Column count is taken from the first row's cell count.
                // If the table has no rows, column count is zero.
                int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

                Console.WriteLine($"Table found: Rows = {rowCount}, Columns = {columnCount}");
            }

            // Save the document (optional, as we only inspected it).
            // Replace "Output.docm" with the desired output path.
            doc.Save("Output.docm");
        }
    }
}
