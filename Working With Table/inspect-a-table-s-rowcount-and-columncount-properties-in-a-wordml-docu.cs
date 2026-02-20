using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load an existing WORDML (or any supported) document.
        Document doc = new Document("Input.docx"); // replace with actual path

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Row count is the number of Row objects in the table.
            int rowCount = table.Rows.Count;

            // Column count is taken from the first row's cell count.
            // If the table has no rows, column count is zero.
            int columnCount = rowCount > 0 ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (unchanged) to demonstrate the required save lifecycle step.
        doc.Save("Output.docx"); // replace with desired output path
    }
}
