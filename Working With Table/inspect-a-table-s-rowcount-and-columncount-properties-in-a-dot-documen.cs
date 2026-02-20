using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOT (template) document.
        Document doc = new Document("Template.dot");

        // Get the collection of tables in the first section's body.
        TableCollection tables = doc.FirstSection.Body.Tables;

        // Iterate through each table and output its row and column counts.
        for (int i = 0; i < tables.Count; i++)
        {
            Table table = tables[i];

            // Number of rows in the table.
            int rowCount = table.Rows.Count;

            // Number of columns – assume a uniform table and use the first row's cell count.
            int columnCount = table.FirstRow?.Cells.Count ?? 0;

            Console.WriteLine($"Table {i}: Rows = {rowCount}, Columns = {columnCount}");
        }

        // Optionally save the document after inspection (no changes made).
        doc.Save("Template_Inspected.dot");
    }
}
