using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the PDF document.
        Document doc = new Document("input.pdf");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];
            int rowCount = table.Rows.Count;
            int columnCount = table.FirstRow != null ? table.FirstRow.Cells.Count : 0;

            Console.WriteLine($"Table {i + 1}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Save the document (unchanged) if needed.
        doc.Save("output.pdf");
    }
}
